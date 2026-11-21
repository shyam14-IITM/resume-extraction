"""Resume folder tree -> one CSV. OCR and section splitting are local; compacting long cells uses an
optional chain of LLMs with automatic fallback.

    python extract.py [folder] [-o resumes.csv] [--compact gemini:MODEL,openai:MODEL,local:MODEL]
                      [--no-of-words 100] [--dpi 300] [--workers 4] [--lang eng]

Settings can live in a .env file (see .env.example): REAL_PATH (used when no folder is given),
COMPACT, NO_OF_WORDS, API keys. Real environment variables and command-line flags win over .env.

The folder is walked depth-first (year > Dir > "IT Software"/"Core" > 1.1, 1.2 ... > resumes): go as deep
as possible, process the resumes found, come back up, move to the next directory. The path is recorded
in the year/dir/category/batch columns.

Per page: use the PDF text layer if it has one, otherwise OCR it with Tesseract.
DOCX files are read from their XML. Text is then split into sections by matching
header lines (fuzzy, so OCR noise is tolerated) and written as one row per resume.
Cells with more than --no-of-words words are compacted to fewer than that (the untouched
text is also written to <out>_full.csv).

Needs:  pip install pymupdf pytesseract pillow   +   the Tesseract program
        (Windows: winget install UB-Mannheim.TesseractOCR)
"""
import argparse, csv, difflib, hashlib, io, json, os, re, shutil, sys, time, urllib.error, urllib.request, zipfile
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pymupdf
import pytesseract
from PIL import Image


def load_env():
    """Read KEY=VALUE lines from .env (next to this script, then the current folder) into os.environ.
    Variables that are already set win. Comment lines start with #."""
    for d in (Path(__file__).parent, Path.cwd()):
        f = d / ".env"
        if f.is_file():
            for line in f.read_text(encoding="utf-8-sig").splitlines():
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    k, v = line.split("=", 1)
                    os.environ.setdefault(k.strip(), v.strip().strip("\"'"))


load_env()
REAL_PATH = os.environ.get("REAL_PATH", "")      # root folder for the final run
NO_OF_WORDS = int(os.environ.get("NO_OF_WORDS") or 100)  # noOfWords: cells with MORE words than this are
                                                         # compacted to FEWER than this
COMPACT = os.environ.get("COMPACT", "extractive")        # model chain, see build_chain()

# canonical column -> header phrases (whole-line match after normalising)
SYNONYMS = {
    "education": ["education", "educational background", "education and scholastic achievements",
                  "academic qualifications", "academics"],
    "experience": ["professional experience", "work experience", "experience", "internships",
                   "internship", "internship experience", "industrial experience",
                   "work and internship experience"],
    "projects": ["projects", "academic projects", "research projects", "personal projects",
                 "key projects", "technical projects", "other projects", "course projects",
                 "projects and research", "research experience"],
    "positions": ["positions of responsibility", "position of responsibility", "responsibilities",
                  "leadership", "leadership and responsibilities"],
    "achievements": ["scholastic achievements", "achievements", "academic achievements",
                     "awards", "awards and achievements", "honours and awards"],
    "skills": ["skills", "technical skills", "skills and tools", "technical proficiency",
               "languages tools and software", "programming skills", "relevant coursework and skills"],
    "courses": ["courses", "course work", "coursework", "relevant coursework", "relevant courses",
                "key courses", "mtech course work", "undergraduate course work"],
    "extracurricular": ["extra curricular", "extracurricular", "extra curricular activities",
                        "extracurricular activities", "co curricular activities", "hobbies",
                        "extra curricular activities and volunteering", "volunteering"],
    "certifications": ["certifications", "certificates", "courses and certifications"],
}
COLUMNS = list(SYNONYMS) + ["unlabeled"]
BULLETS = "\u2022\u25cf\u25aa\u25a0\u2043\u2219\u25e6\ufffd\uf0b7\uf0a7\u2023"


def norm(line):
    s = line.strip().strip(BULLETS + " \t").lower().replace("&amp;", "and").replace("&", "and")
    s = s.replace("-", " ")
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9 ]", "", s)).strip()


PHRASE = {norm(p): col for col, ps in SYNONYMS.items() for p in ps}


# For ALL-CAPS header lines not in SYNONYMS: the keyword that appears first decides the column
# ("PROFESSIONAL EXPERIENCE AND PROJECTS" -> experience). Stems are matched as substrings.
KEYWORDS = [("education", "education"), ("experience", "experience"), ("internship", "experience"),
            ("endeavour", "experience"), ("employment", "experience"), ("project", "projects"),
            ("responsibilit", "positions"), ("position", "positions"), ("leadership", "positions"),
            ("achievement", "achievements"), ("award", "achievements"), ("honour", "achievements"),
            ("certification", "certifications"), ("coursework", "courses"), ("course", "courses"),
            ("skill", "skills"), ("proficiency", "skills"), ("curricular", "extracurricular"),
            ("hobbies", "extracurricular"), ("volunteer", "extracurricular"), ("activities", "extracurricular")]


def header_of(line):
    s = line.strip().strip(BULLETS + " \t")
    n = norm(line)
    if len(s) > 70 or len(n) < 4:
        return None
    if n in PHRASE:
        return PHRASE[n]
    close = difflib.get_close_matches(n, PHRASE, n=1, cutoff=0.88)  # tolerate OCR typos
    if close:
        return PHRASE[close[0]]
    letters = [c for c in s if c.isalpha()]
    if len(n.split()) <= 8 and len(letters) >= 4 and sum(c.isupper() for c in letters) / len(letters) > 0.85:
        hits = [(n.find(k), col) for k, col in KEYWORDS if k in n]
        return min(hits)[1] if hits else None
    return None


def clean(text):
    text = text.replace("&amp;", "&").replace("&quot;", '"').replace("&apos;", "'")
    text = re.sub(r"[\ufffd\uf0b7\uf0a7\u25cf]", "\u2022", text)
    text = re.sub(r"[ \t]+", " ", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def sectionize(text):
    sections = {c: [] for c in COLUMNS}
    current = "unlabeled"
    for line in text.splitlines():
        h = header_of(line)
        if h:
            current = h
        else:
            sections[current].append(line)
    return {c: clean("\n".join(v)) for c, v in sections.items()}


def first(pattern, text, flags=re.I):
    m = re.search(pattern, text, flags)
    return m.group(1) if m else ""


# ---------- reading files ----------

def sniff(path):
    """Real type from content, since extensions here are unreliable (.pdf_, .pdf ...)."""
    with open(path, "rb") as f:
        head = f.read(4)
    if head == b"%PDF":
        return "pdf"
    if head[:2] == b"PK":
        try:
            if "word/document.xml" in zipfile.ZipFile(path).namelist():
                return "docx"
        except zipfile.BadZipFile:
            pass
    return None


def walk_files(root):
    """Depth-first: descend into each sub-directory as far as it goes (natural order, so 1.2 comes
    before 1.10), yield the files found, then come back up and continue with the next directory."""
    def natural(p):
        return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", p.name)]
    entries = sorted(root.iterdir(), key=natural)
    for d in (e for e in entries if e.is_dir()):
        yield from walk_files(d)
    yield from (e for e in entries if e.is_file())


def normalize_names(files):
    """Rename PDF/DOCX files in place so the extension matches the content
    (x.pdf_ / x.pdf / 'x. pdf' -> x.pdf, x.pdf_.docx -> x.docx). Other files are untouched.
    Returns the file list with the new names."""
    result = []
    for p in files:
        kind = sniff(p)
        if not kind:
            result.append(p)
            continue
        base = p.name
        while True:  # peel trailing junk: spaces, dots, underscores, stray "pdf"/"docx"
            new = re.sub(r"[\s._]*(?:pdf|docx)?[\s._]*$", "", base, flags=re.I)
            if new == base:
                break
            base = new
        target = p.with_name(f"{base}.{kind}")
        if target == p:
            result.append(p)
        elif target.exists():
            print(f"  rename skipped (target exists): {p.name} -> {target.name}")
            result.append(p)
        else:
            p.rename(target)
            print(f"  renamed: {p.name} -> {target.name}")
            result.append(target)
    return result


# ---------- compacting long cells ----------
# COMPACT / --compact is a chain "provider:model,provider:model,...". For every long cell the entries
# are tried left to right; on a rate limit or other error the next entry is used. If every entry is
# unusable the cell falls back to extractive compaction (no model), so a run never stops half-way.

DEFAULT_MODEL = {"local": "qwen2.5:3b", "openai": "gpt-4o-mini", "gemini": "gemini-2.5-flash"}
ENV_KEY = {"openai": "OPENAI_API_KEY", "gemini": "GEMINI_API_KEY"}
MAX_STRIKES = 5     # consecutive rate-limit/transient failures before an entry is dropped for the run
MAX_COOLDOWN = 120  # seconds


def wc(text):
    return len(text.split())


def extractive_compact(text, n):
    """No model: keep the most informative lines (ones with numbers/metrics first), in original
    order, within n-5 words. A single over-long line is cut at the budget."""
    limit = max(1, n - 5)
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    order = sorted(range(len(lines)), key=lambda i: (-(2 * bool(re.search(r"\d", lines[i])) + (i == 0)), i))
    keep, used = set(), 0
    for i in order:
        if used + wc(lines[i]) <= limit:
            keep.add(i)
            used += wc(lines[i])
    out = [lines[i] for i in sorted(keep)]
    return "\n".join(out) if out else " ".join(lines[0].split()[:limit]) + " ..."


class LLMError(Exception):
    """kind: rate (cool down, retry later) | transient (server/network, cool down) |
    fatal (bad key, unknown model, no quota: drop the entry) | skip (this request only)."""

    def __init__(self, msg, kind, retry_after=60):
        super().__init__(msg)
        self.kind, self.retry_after = kind, min(retry_after, MAX_COOLDOWN)


def _post(url, payload, headers=None):
    req = urllib.request.Request(url, method="POST", data=json.dumps(payload).encode(),
                                 headers={"Content-Type": "application/json", **(headers or {})})
    try:
        with urllib.request.urlopen(req, timeout=300) as r:
            return json.load(r)
    except urllib.error.HTTPError as e:
        body = e.read()[:300].decode("utf-8", "replace")
        if e.code == 429 and "insufficient_quota" in body:
            kind = "fatal"
        elif e.code == 429:
            kind = "rate"
        elif e.code in (401, 403, 404):
            kind = "fatal"
        elif e.code >= 500:
            kind = "transient"
        else:
            kind = "skip"
        ra = e.headers.get("Retry-After", "")
        raise LLMError(f"HTTP {e.code} {' '.join(body.split())[:120]}", kind,
                       int(ra) if ra.isdigit() else 60) from None
    except (urllib.error.URLError, TimeoutError, OSError) as e:
        raise LLMError(str(e), "transient") from None


def ask_local(prompt, model, args):
    return _post(args.ollama_url.rstrip("/") + "/api/generate",
                 {"model": model, "prompt": prompt, "stream": False,
                  "options": {"temperature": 0, "num_predict": 220, "num_ctx": 4096}})["response"]


def ask_openai(prompt, model, args):
    r = _post("https://api.openai.com/v1/chat/completions",
              {"model": model, "messages": [{"role": "user", "content": prompt}]},
              {"Authorization": "Bearer " + os.environ["OPENAI_API_KEY"]})
    return r["choices"][0]["message"]["content"]


def ask_gemini(prompt, model, args):
    r = _post(f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent",
              {"contents": [{"parts": [{"text": prompt}]}], "generationConfig": {"temperature": 0}},
              {"x-goog-api-key": os.environ["GEMINI_API_KEY"]})
    return r["candidates"][0]["content"]["parts"][0]["text"]


ASK = {"local": ask_local, "openai": ask_openai, "gemini": ask_gemini}


class Entry:
    def __init__(self, provider, model):
        self.provider, self.model = provider, model
        self.name = f"{provider}:{model}"
        self.until, self.strikes, self.dead = 0.0, 0, False


def build_chain(args):
    """Parse the chain spec, drop entries that can't work (missing key / Ollama down), exit if none are left."""
    spec = [s.strip() for s in args.compact.split(",") if s.strip() and s.strip().lower() != "extractive"]
    chain = []
    for item in spec:
        provider, _, model = item.partition(":")
        provider = provider.lower()
        if provider not in ASK:
            sys.exit(f"Unknown provider '{provider}' in --compact/COMPACT (use gemini, openai, local, or extractive).")
        entry = Entry(provider, model or DEFAULT_MODEL[provider])
        if provider in ENV_KEY and not os.environ.get(ENV_KEY[provider]):
            print(f"  skipping {entry.name}: {ENV_KEY[provider]} is not set")
            continue
        if provider == "local":
            try:
                urllib.request.urlopen(args.ollama_url.rstrip("/") + "/api/tags", timeout=5)
            except Exception:
                print(f"  skipping {entry.name}: Ollama not reachable at {args.ollama_url}")
                continue
        chain.append(entry)
    if spec and not chain:
        sys.exit("None of the models in the compact chain are usable. Fill in the keys in .env, start Ollama, "
                 "or use --compact extractive.")
    return chain


class Compactor:
    """compactor(text) -> text with fewer than `no_of_words` words. LLM results are cached on disk
    (keyed by limit + text) so an interrupted run resumes where it stopped and costs nothing twice."""

    def __init__(self, args, cache_path):
        self.n, self.entries, self.path, self.args_for_ask = args.no_of_words, args.chain, cache_path, args
        self.stats = Counter()
        try:
            self.cache = json.loads(cache_path.read_text(encoding="utf-8")) if self.entries else {}
        except (OSError, ValueError):
            self.cache = {}

    def __call__(self, text):
        if wc(text) <= self.n:
            return text
        if not self.entries:
            self.stats["extractive"] += 1
            return extractive_compact(text, self.n)
        key = hashlib.sha1(f"{self.n}\0{text}".encode()).hexdigest()
        if key in self.cache:
            self.stats["cache"] += 1
            return self.cache[key]
        out = self._ask_chain(text)
        if out is None:
            self.stats["extractive (all models failed)"] += 1
            return extractive_compact(text, self.n)
        self.cache[key] = out
        self.path.write_text(json.dumps(self.cache), encoding="utf-8")
        return out

    def _ask_chain(self, text):
        while True:
            live = [e for e in self.entries if not e.dead]
            if not live:
                return None
            ready = [e for e in live if e.until <= time.time()]
            if not ready:  # everything left is cooling down after a rate limit: wait for the first to recover
                wait = max(1.0, min(e.until for e in live) - time.time())
                print(f"  all models busy; waiting {wait:.0f}s", flush=True)
                time.sleep(wait)
                continue
            errored = False
            for e in ready:
                try:
                    out = self._try(e, text)
                except LLMError as err:
                    errored = True
                    self._fail(e, err)
                    continue
                e.strikes = 0
                if out:
                    self.stats[e.name] += 1
                    return out
            if not errored:  # models answered but never got under the limit
                return None

    def _try(self, e, text):
        # models overshoot the word target, so ask for ~55% of the limit, then ~35%
        for target in (int(self.n * 0.55), int(self.n * 0.35)):
            prompt = (f"Summarise this resume section in at most {max(target, 5)} words, as short plain lines. "
                      "Keep only facts that are in the text: names, numbers, technologies, dates. "
                      "Do not add anything new. Output only the summary.\n\n" + text)
            try:
                out = ASK[e.provider](prompt, e.model, self.args_for_ask).strip()
            except (KeyError, IndexError, TypeError):
                raise LLMError("unexpected response shape", "skip") from None
            if 0 < wc(out) < self.n:
                return out
        return None

    def _fail(self, e, err):
        print(f"  {e.name}: {err.kind}: {err}", flush=True)
        if err.kind == "fatal":
            e.dead = True
        elif err.kind in ("rate", "transient"):
            e.strikes += 1
            e.until = time.time() + (err.retry_after if err.kind == "rate" else 30)
            if e.strikes >= MAX_STRIKES:
                e.dead = True
                print(f"  {e.name}: failing repeatedly, dropped for this run", flush=True)
        if e.dead and err.kind == "fatal":
            print(f"  {e.name}: dropped for this run", flush=True)

    def summary(self):
        return ", ".join(f"{k}: {v}" for k, v in self.stats.items()) or "no cells needed compacting"


def docx_text(path):
    xml = zipfile.ZipFile(path).read("word/document.xml").decode("utf-8")
    paras = re.findall(r"<w:p[ >].*?</w:p>", xml, flags=re.S)
    lines = ["".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", p)) for p in paras]
    return "\n".join(l for l in lines if l.strip())


def ocr_page(page, dpi, lang):
    pix = page.get_pixmap(dpi=dpi)
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    # psm 3 = automatic page segmentation (detects columns); keep spacing between words
    return pytesseract.image_to_string(img, lang=lang, config="--psm 3 -c preserve_interword_spaces=1")


def pdf_text(path, dpi, lang):
    """-> (text, how) where how says which pages were OCR'd."""
    parts, ocr_pages = [], []
    with pymupdf.open(path) as doc:
        for i, page in enumerate(doc, 1):
            t = page.get_text(sort=True)
            if len(t.strip()) < 30:  # no real text layer on this page
                t = ocr_page(page, dpi, lang)
                ocr_pages.append(i)
            parts.append(t)
        n = len(doc)
    how = "text layer" if not ocr_pages else ("OCR" if len(ocr_pages) == n else f"OCR p{ocr_pages}+text")
    return "\n".join(parts), how


def process(path, args, root):
    kind = sniff(path)
    if kind is None:
        return None
    text, how = (docx_text(path), "docx") if kind == "docx" else pdf_text(path, args.dpi, args.lang)
    stem = re.sub(r"(\.pdf.*|\.docx)$", "", path.name, flags=re.I)
    rel = path.parent.relative_to(root).parts  # e.g. (Dir1, "IT Software", "1.1")
    level = lambda i: rel[i] if len(rel) > i else ""
    return {
        "file": path.name,
        "year": root.name,
        "dir": level(0),
        "category": level(1),
        "batch": level(2),
        "path": "/".join((root.name, *rel)),
        "company": re.split(r"[_ ]", stem, maxsplit=1)[0].title(),
        "extraction": how,
        "cgpa": first(r"(?:cgpa|gpa)\W{0,6}(\d{1,2}\.\d{1,2})", text)
                or first(r"\b(\d\.\d{1,2})\s*/\s*10\b", text),
        "emails": ";".join(sorted(set(re.findall(r"[\w.+-]+@[\w-]+\.[\w.-]+", text)))),
        **sectionize(text),
    }


def find_tesseract():
    exe = shutil.which("tesseract")
    if not exe:
        default = Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe")
        exe = str(default) if default.exists() else None
    if not exe:
        sys.exit("Tesseract not found. Install it (Windows: winget install UB-Mannheim.TesseractOCR) "
                 "or add tesseract.exe to PATH.")
    pytesseract.pytesseract.tesseract_cmd = exe


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("folder", type=Path, nargs="?", default=REAL_PATH or None,
                    help="root folder to walk (default: REAL_PATH from .env)")
    ap.add_argument("-o", "--out", type=Path, default=Path("resumes.csv"))
    ap.add_argument("--dpi", type=int, default=300)
    ap.add_argument("--workers", type=int, default=4)
    ap.add_argument("--lang", default="eng")
    ap.add_argument("--compact", default=COMPACT,
                    help="fallback chain for compacting long cells, 'provider:model,provider:model,...' with "
                         "providers gemini, openai, local (Ollama); the next entry is used on a rate limit or "
                         "error. 'extractive' = no model. Default: COMPACT from .env, else extractive. "
                         "Cloud entries send resume text to that service.")
    ap.add_argument("--no-of-words", "--max-words", dest="no_of_words", type=int, default=NO_OF_WORDS,
                    help="cells with more words than this are compacted to fewer than this "
                         f"(default {NO_OF_WORDS})")
    ap.add_argument("--ollama-url", default=os.environ.get("OLLAMA_URL", "http://localhost:11434"))
    args = ap.parse_args()
    if not args.folder:
        sys.exit("No folder given. Pass one on the command line or set REAL_PATH in .env.")
    args.folder = Path(args.folder)
    if not args.folder.is_dir():
        sys.exit(f"Not a folder: {args.folder}")
    args.chain = build_chain(args)
    find_tesseract()

    out_files = {args.out.resolve()}
    files = [p for p in walk_files(args.folder) if p.resolve() not in out_files]
    files = normalize_names(files)
    rows, skipped = [], []

    def work(p):
        try:
            r = process(p, args, args.folder)
            print(f"  {'done' if r else 'skip'}: {p.name}" + (f"  [{r['extraction']}]" if r else "  (not a PDF/DOCX)"),
                  flush=True)
            return p, r, None
        except Exception as e:
            print(f"  FAIL: {p.name}: {e}", flush=True)
            return p, None, e

    with ThreadPoolExecutor(args.workers) as ex:
        for p, r, err in ex.map(work, files):
            if r:
                rows.append(r)
            elif err:
                skipped.append((p.name, f"error: {err}"))

    fields = ["file", "year", "dir", "category", "batch", "path", "company", "extraction", "cgpa", "emails"] + COLUMNS
    def write(path, data):
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=fields)
            w.writeheader()
            w.writerows(data)

    full_path = args.out.with_name(args.out.stem + "_full" + args.out.suffix)
    write(full_path, rows)  # untouched text, so nothing is lost by shortening
    compact = Compactor(args, args.out.with_name(args.out.stem + "_llm_cache.json"))
    short = []
    for n, r in enumerate(rows, 1):  # sequential: calls are the bottleneck (and local LLMs don't parallelise)
        short.append({k: compact(v) if k in COLUMNS else v for k, v in r.items()})
        if args.chain:
            print(f"  compacted {n}/{len(rows)}: {r['file']}", flush=True)
    write(args.out, short)
    print(f"\nwrote {len(rows)} rows -> {args.out} (fields < {args.no_of_words} words) and {full_path} (full text)")
    print(f"compacted cells by: {compact.summary()}")
    for name, why in skipped:
        print(f"  failed: {name} ({why})")


if __name__ == "__main__":
    main()

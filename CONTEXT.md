# Resume Scraping: Context

## Goal
Turn hundreds of resume files (PDF/DOCX, varying layouts) into **one CSV** with one row per resume.
Long sections (projects, experience, etc.) go in text columns. **No external/cloud LLMs**; everything runs locally and free.

## Data (`2024-25/`, `2025-26/` currently empty)
- 28 files named `<Company>_<Role> (n).pdf` (extensions are messy: `.pdf_`, `.pdf `, plus 2 `.docx`). Detect type by content, not extension.
- Names and contact details are redacted (black boxes). Use the filename as the record ID.
- **Only 10 of 28 files have a text layer.** 2 DOCX have the full text. In 8 PDFs, page 1 is an image and only page 2 has text. The other 18 are image-only.
- Layouts vary: single column with tables, and dense two-column with a left sidebar for company/role.
- Section titles vary: "Projects" / "Academic Projects" / "Research Projects"; "Internship" / "Professional Experience"; "Positions of Responsibility"; "Scholastic Achievements"; etc.

## Current state
`extract.py` (regex only, no LLM):
- Reads text with `pdftotext` (poppler) and the DOCX XML. Image-only files are skipped and listed.
- Splits text into sections by matching short header lines against a synonym table (`SYNONYMS` at the top of the script).
- Writes `resumes.csv` with `file, year_folder, company, text_coverage, cgpa, gate_or_jee_rank, emails, education, experience, projects, positions, achievements, skills, courses, extracurricular, certifications, unlabeled`.
- Run: `python extract.py 2024-25 resumes.csv`
- Known limits: DOCX text comes out in XML order, so sidebar and column text can land in the wrong section. `unlabeled` holds page-1 content that is missing.

## Next step: local OCR for image-only pages
Machine: **Lenovo LOQ, RTX 3050 6 GB, 16 GB RAM, i5** (no GPU was visible from the earlier session, so check `nvidia-smi` first).

Plan (two stages, both local and free):
1. Render each page to an image at about 200 dpi and read it into markdown/text with a local model that keeps column order.
2. Run the same regex section splitter on that text and write the same CSV.

| Model | Fits 6 GB? | Notes |
|---|---|---|
| **PaddleOCR-VL (about 0.9B)** | Yes | Document parser, handles tables/columns. Small and fast. Recommended. |
| **Qwen2.5-VL 3B** (Ollama / llama.cpp, 4-bit) | Yes (about 3 GB) | General vision model, copes with odd layouts. Recommended. |
| Qwen2.5-VL 7B / MiniCPM-V (about 8B), Q4 | Barely | More accurate, tight on memory. |
| olmOCR / GOT-OCR2 (about 0.6B) | Yes | Fast page-to-markdown. |
| **RapidOCR** (ONNX) or **Surya** | Yes (CPU or GPU) | Classical OCR, lightest option. Weaker on sidebar layouts. |

Rough speed: about 2–15 s per page, so hundreds of resumes take under an hour.

## Open items
- Pick the OCR model (see above) and wire it into `extract.py` as a fallback for pages with no text layer.
- Spot-check OCR output (CGPA, ranks, `0` vs `O`), especially the sidebar-layout resumes.
- Add a per-file quality report (sections found, empty columns).

# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**oncquest-pdf-test** is a Streamlit-based universal PDF processor that extracts structured data from PDFs and converts between multiple formats (JSON, HTML, Word .docx).

**Core workflow:**
- PDF → Extraction (structured JSON with tables, images, metadata)
- JSON + theme.json → Multiple output formats (HTML, DOCX, PDF preview)
- Supports lab reports, invoices, and general document processing

## Setup & Development Environment

### Initial Setup (Isolated venv)

```bash
python3 -m venv .venv-oncquest
.venv-oncquest/bin/pip install -r requirements.txt
.venv-oncquest/bin/playwright install chromium
```

### Activate Environment

```bash
source .venv-oncquest/bin/activate
```

### Run Streamlit App

```bash
streamlit run app.py
```

The app runs on `http://localhost:8501` by default.

### Teardown (Clean remove venv)

```bash
./teardown.sh
# OR manual: rm -rf .venv-oncquest && playwright uninstall --all
```

## Core Architecture

### Main Modules

1. **`app.py`** — Streamlit UI entry point
   - Page config, file upload handlers (PDF/JSON)
   - Tabs for extraction, conversion, preview modes
   - Theme.json config loader (from repo root)

2. **`extractor.py`** — PDF → JSON extraction engine
   - **`extract_report_data(pdf_path)`** — Main entry point
   - Extracts: key-value pairs, tables (with headers/grids), images (Base64), page layout
   - Outputs: structured JSON with bounding boxes and typography metadata
   - Handles SNG/Genelab lab report name sanitization

3. **`converter.py`** — Rendering & format conversion module
   - **`render_exact_pdf_layout_html(data, theme)`** — Preserves 100% exact PDF layout (absolute positioning)
   - **`generate_dynamic_template_html(data, theme)`** — Responsive HTML flow template
   - **`convert_json_to_docx(data, theme)`** — Direct JSON → .docx export
   - **`convert_pdf_to_word(pdf_path)`** — PDF → .docx (multiple methods, fallback-based)
   - **`render_html_to_pdf_and_preview(html)`** — HTML → PDF via Playwright
   - Text/lab name replacements: `replace_test_name_in_html()`, `replace_sng_gen_lab()`

4. **`convert.py`** — Standalone conversion CLI utility (legacy/supplementary)

5. **`replace_test_names.py`** — Batch renaming utility for test names in JSON files

### Data Flow

```
PDF file (app.py upload)
    ↓
extract_report_data() [extractor.py]
    ↓
Structured JSON: {pages: [...], tables: [...], images: [...], metadata: {...}}
    ↓
(saved to extracted_jsons/ and optionally imported back)
    ↓
Theme.json (styling config)
    ↓
render_exact_pdf_layout_html() OR generate_dynamic_template_html() [converter.py]
    ↓
HTML preview (Streamlit display)
    ↓
convert_json_to_docx() OR render_html_to_pdf_and_preview()
    ↓
.docx or .pdf output (download via app.py)
```

### Theme & Styling

- **`theme.json`** — Central style config (colors, fonts, margins, table styles, etc.)
- Loaded by `app.py` on startup; passed to all converter functions
- If missing, converters use built-in defaults

### Output Directories

- **`extracted_jsons/`** — Saved JSON extractions from PDF uploads
- **`output/`** — Generated .docx, .pdf files (auto-created)
- **`scratch/`** — Ad-hoc analysis scripts (not part of core workflow)
- **`outsourcing pdf/`** — Separate test/demo PDFs

## Key Patterns

### Lab Report Sanitization

Both `extractor.py` and `converter.py` include text-replacement functions to sanitize sensitive lab names:
- `replace_sng_gen_lab()` — Replaces "SNG Gene's Lab" variants with "Laboratory"
- `replace_test_name_in_structure()` — Replaces specific test names ("Liquidseq...", "Brainseq...") with "TEST NAME"

These run on JSON structures and HTML. Applied automatically in extraction and conversion pipelines.

### Fallback Conversion Methods

`converter.py` offers multiple PDF→DOCX strategies:
1. **Method-1**: Direct python-docx rendering from extracted JSON
2. **Method-2**: pdf2docx library (lossy but fast)
3. **Method-3**: HTML → DOCX via Playwright + htmldocx (optional, unmaintained)

App chooses method based on availability; all write to `.docx` format.

## Dependencies

- **streamlit** — UI framework
- **pymupdf** — PDF extraction, rendering, layout analysis
- **python-docx** — Word document generation
- **pillow** — Image processing
- **beautifulsoup4** — HTML parsing
- **playwright** — Headless browser (HTML→PDF, chromium installation required)
- **pdf2docx** — Alternative PDF→DOCX converter

See `requirements.txt` for version pins.

## Common Commands

```bash
# Extract JSON from a PDF (via Streamlit UI)
streamlit run app.py
# → Upload PDF → "Extract & Preview JSON" tab

# Batch-process: rename test names in extracted JSONs
python3 replace_test_names.py <input_json_dir> <output_json_dir>

# Ad-hoc inspection
python3 scratch/analyze_json.py <extracted.json>
python3 scratch/find_term.py <extracted.json> <search_term>
```

## Testing & Debugging

- **Scratch folder** (`scratch/`) contains exploration scripts:
  - `analyze_json.py` — Inspect JSON structure
  - `check_docx.py` — Verify generated DOCX files
  - `verify_widths.py` — Debug table/column widths
  - `inspect_table.py` — Deep-dive table extraction

- **Playwright** — Required for HTML→PDF. Install browsers with:
  ```bash
  playwright install chromium
  ```

## Git Workflow

- **Current branch**: `correct_JSON` (ongoing work)
- **Main branch**: `main`
- Commits include infrastructure setup (venv, requirements, teardown script)

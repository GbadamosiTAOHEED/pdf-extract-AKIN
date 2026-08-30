# PDF Extraction Suite Pro

A Streamlit app that OCRs a scanned PDF directory (e.g. a health facility or
business listing), cleans up the text, and exports a structured Word
document and/or Excel spreadsheet.

## Setup

1. **Install Python packages**
   ```bash
   pip install -r requirements.txt
   ```

2. **Install the two external tools** (these are *not* pip packages):
   - **Tesseract OCR** — https://github.com/UB-Mannheim/tesseract/wiki (Windows installer)
     macOS: `brew install tesseract` · Linux: `sudo apt install tesseract-ocr`
   - **Poppler** (used by `pdf2image`) — https://github.com/oschwartz10612/poppler-windows/releases (Windows)
     macOS: `brew install poppler` · Linux: `sudo apt install poppler-utils`

   On macOS/Linux, once installed via a package manager, both tools are
   already on your `PATH` — you can skip step 3 and leave the sidebar path
   fields blank.

3. **Point the app at your local install (Windows, or a custom install location)**
   Copy `.env.example` to `.env` and fill in your paths:
   ```
   TESSERACT_CMD=C:\Program Files\Tesseract-OCR\tesseract.exe
   POPPLER_PATH=C:\path\to\poppler-xx.xx.x\Library\bin
   ```
   `.env` is git-ignored, so your personal paths never get committed. If you
   skip this, you can still type the paths directly into the app's sidebar
   each time you run it.

## Run

```bash
streamlit run pdf_extraction_suite.py
```

> Run it with `streamlit run`, **not** `python pdf_extraction_suite.py` —
> launching it as a plain script produces harmless-but-noisy
> `missing ScriptRunContext` warnings in the terminal because Streamlit
> isn't running inside its own server in that mode.

This opens the app in your browser at `http://localhost:8501`. Upload a
scanned PDF, adjust the DPI/quality settings in the sidebar if needed, and
click **Start processing**.

## Notes on speed

- **DPI 200** (default) is enough for most printed directories; only raise
  it if the source text is genuinely tiny.
- **Preprocessing quality**: leave on *Fast* or *Balanced* for clean,
  digitally printed pages. *Max accuracy* adds a slow denoise pass — only
  worth it on noisy photocopies.
- Use **"Only process the first N pages"** to test your settings on a big
  file before committing to a full run.

## Sharing this with your team

- This repo is safe to share as-is — no personal paths or uploaded data are
  committed (see `.gitignore`).
- Each teammate installs Tesseract/Poppler locally and sets their own
  `.env` (or just fills in the sidebar).
- Because this may process health-facility or organizational directory
  data, keep the repository **private** on GitHub/GitLab rather than
  public, and avoid committing any real output files (`.docx`/`.xlsx`) —
  the `.gitignore` already excludes them, but double-check before pushing.

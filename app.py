"""
PDF Extraction Suite AKIN
-------------------------
A Streamlit app that OCRs a scanned PDF (e.g. a business/health-facility
directory), cleans the text, and exports it as a structured Word document
and/or Excel spreadsheet.

RUN:
    pip install streamlit pdf2image opencv-python-headless numpy pandas \
                pytesseract python-docx openpyxl pillow
    streamlit run pdf_extraction_suite.py

REQUIREMENTS ON YOUR MACHINE (these are NOT pip packages):
    1. Tesseract OCR engine  -> https://github.com/UB-Mannheim/tesseract/wiki
    2. Poppler (for pdf2image) -> https://github.com/oschwartz10612/poppler-windows/releases
       (On macOS: `brew install poppler`. On Linux: `sudo apt install poppler-utils`
        and pytesseract's `apt install tesseract-ocr` — in which case you can
        leave the sidebar paths blank, since both are already on PATH.)

Fill in the Tesseract/Poppler paths once in the sidebar (or hardcode your
defaults below) and the app remembers nothing else — everything happens
in-memory, nothing is written to disk.
"""

import io
import os
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

import cv2
import numpy as np
import pandas as pd
import pytesseract
import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from pdf2image import convert_from_bytes, pdfinfo_from_bytes
from PIL import Image

try:
    from dotenv import load_dotenv
    load_dotenv()  # loads a local, git-ignored .env file if present — see .env.example
except ImportError:
    pass  # python-dotenv is optional; sidebar fields still work without it

# --------------------------------------------------------------------------
# 1. Page configuration
# --------------------------------------------------------------------------
st.set_page_config(page_title="PDF Extraction Suite AKIN", page_icon="5891226990491143527_121.jpg", layout="centered")

# --------------------------------------------------------------------------
# 2. Default local paths (edit these to your own machine, or just fill them
#    in the sidebar at runtime — sidebar values always win).
# --------------------------------------------------------------------------
# --------------------------------------------------------------------------
# 2. Default local paths. These are picked up from environment variables
#    (or a local, git-ignored `.env` file — see .env.example) so nobody's
#    personal folder structure ends up committed to the repo. Anyone on the
#    team can still just type their own path into the sidebar at runtime
#    without touching this file at all.
# --------------------------------------------------------------------------
DEFAULT_TESSERACT_CMD = os.environ.get("TESSERACT_CMD", r"C:\Program Files\Tesseract-OCR\tesseract.exe")
DEFAULT_POPPLER_PATH = os.environ.get("POPPLER_PATH", "")

ADDRESS_KEYWORD_PATTERN = re.compile(
    r'\b(?:street|st\.?|road|rd\.?|avenue|ave\.?|plot|close|crescent|way|'
    r'estate|layout|opposite|opp\.?|beside|junction|roundabout|market|'
    r'lagos|abuja|ibadan|oyo\s+state|state|p\.?\s?o\.?\s?box|gra)\b',
    re.IGNORECASE,
)

PHONE_REGEX = re.compile(r'(\+?\d[\d\s\-\(\)]{6,}\d)')
EMAIL_WEB_REGEX = re.compile(
    r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    r'|www\.[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    r'|https?://[^\s]+)',
    re.IGNORECASE,
)


# --------------------------------------------------------------------------
# 3. Image preprocessing (OpenCV)
# --------------------------------------------------------------------------
def preprocess_image(pil_img: Image.Image, mode: str = "balanced") -> Image.Image:
    """Denoise, boost contrast, and binarize a scanned page for OCR.

    mode:
      "fast"      - adaptive threshold only. Best for clean, digitally
                     printed pages (most factbooks/directories). Several
                     times faster than the other two modes.
      "balanced"  - + CLAHE contrast boost (cheap, helps faded print).
      "max"       - + bilateral filter denoise (slow — only worth it on
                     genuinely noisy/speckled photocopies).
    """
    # Image already arrives as single-channel grayscale (we render the PDF
    # with grayscale=True), so there's no RGB->gray conversion to pay for.
    gray = np.array(pil_img)
    if gray.ndim == 3:
        gray = cv2.cvtColor(gray, cv2.COLOR_RGB2GRAY)

    work = gray
    if mode == "max":
        # Preserve edges while removing scanner speckle noise — the
        # slowest step in the whole pipeline, so it's opt-in only.
        work = cv2.bilateralFilter(work, d=9, sigmaColor=75, sigmaSpace=75)

    if mode in ("balanced", "max"):
        clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
        work = clahe.apply(work)

    # Adaptive threshold copes with uneven scan lighting far better than a
    # single global threshold, and is cheap regardless of mode.
    binary = cv2.adaptiveThreshold(
        work, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY, 21, 11,
    )
    return Image.fromarray(binary)


def ocr_single_page(args, tesseract_cmd: str, mode: str = "balanced"):
    """OCR one page image and return (page_num, line) tuples."""
    page_num, img = args
    pytesseract.pytesseract.tesseract_cmd = tesseract_cmd
    lines = []
    try:
        clean_img = preprocess_image(img, mode=mode)
        # PSM 4 (single column of variable-sized text) tracks a directory's
        # vertical flow more faithfully than PSM 6.
        config = r"--oem 3 --psm 4 -c preserve_interword_spaces=1"
        text = pytesseract.image_to_string(clean_img, config=config)
        for line in text.splitlines():
            stripped = line.strip()
            if stripped:
                lines.append((page_num, stripped))
    except Exception as exc:  # keep going even if one page fails
        lines.append((page_num, f"[OCR ERROR ON THIS PAGE: {exc}]"))
    return lines


# --------------------------------------------------------------------------
# 4. Text cleanup
# --------------------------------------------------------------------------
def fix_ocr_typos(text: str) -> str:
    """Normalize the most common scan artefacts (P.O. Box variants, stray
    symbols Tesseract likes to invent)."""
    # Any of the frequent P.O. Box mis-reads -> "P.O. Box"
    text = re.sub(r'\bP[\.,]?\s?[O0Q][\.,]?\s?Box\b', 'P.O. Box', text, flags=re.IGNORECASE)
    text = re.sub(r'\b[2Z][\.,]?\s?[O0][\.,]?\s?Box\b', 'P.O. Box', text, flags=re.IGNORECASE)
    # Collapse doubled/triple spaces left behind after symbol removal
    text = re.sub(r'[\}\{\|®©~^`]', '', text)
    text = re.sub(r'\s{2,}', ' ', text)
    return text.strip()


def is_category_header(line: str) -> bool:
    """Heuristic: short, ALL-CAPS lines with no phone/email are section
    headers (e.g. 'PHARMACIES', 'HOSPITALS & CLINICS')."""
    letters = re.sub(r'[^A-Za-z]', '', line)
    if not letters:
        return False
    return (
        line.isupper()
        and 3 <= len(line) <= 55
        and not PHONE_REGEX.search(line)
        and "@" not in line
    )


def looks_like_address(line: str) -> bool:
    return bool(ADDRESS_KEYWORD_PATTERN.search(line))


# --------------------------------------------------------------------------
# 5. Structured parsing -> list[dict] (used for both Word and Excel output)
# --------------------------------------------------------------------------
def parse_lines_to_records(all_lines):
    """Turn raw OCR (page, line) tuples into directory-style records:
    Category, Name, Address, Phone, Email/Website."""
    records = []
    current_category = "General"
    i, n = 0, len(all_lines)

    while i < n:
        page_num, raw = all_lines[i]
        line = fix_ocr_typos(raw)

        if not line or line.startswith("[OCR ERROR"):
            i += 1
            continue

        if is_category_header(line):
            current_category = line.title()
            i += 1
            continue

        phones = PHONE_REGEX.findall(line)
        phone_str = ", ".join(p.strip() for p in phones)

        email_match = EMAIL_WEB_REGEX.search(line)
        email_web = email_match.group(0) if email_match else ""

        name = line
        if email_web:
            name = name.replace(email_web, "")
        for p in phones:
            name = name.replace(p, "")
        name = name.strip(" ,;:-")

        # A line with no leftover name text is a continuation of the entry
        # above (a phone/email that landed on its own OCR line), not a new
        # entry — fold it into the previous record and move on. Doing this
        # BEFORE any address lookahead stops a later, unrelated line from
        # being misread as "this orphan line's address".
        if not name and (phone_str or email_web) and records:
            prev = records[-1]
            if phone_str and not prev["Phone"]:
                prev["Phone"] = phone_str
            elif phone_str:
                prev["Phone"] += f", {phone_str}"
            if email_web and not prev["Email / Website"]:
                prev["Email / Website"] = email_web
            i += 1
            continue

        if not name and not phone_str and not email_web:
            i += 1
            continue

        address_parts = []
        # Pull in up to two following lines if they read like an address
        # and don't themselves look like a new directory entry.
        lookahead = 0
        while lookahead < 2 and (i + 1 + lookahead) < n:
            _, next_raw = all_lines[i + 1 + lookahead]
            next_line = fix_ocr_typos(next_raw)
            if next_line and looks_like_address(next_line) and not is_category_header(next_line):
                address_parts.append(next_line)
                lookahead += 1
            else:
                break
        i += lookahead

        records.append({
            "Page": page_num,
            "Category": current_category,
            "Name": name,
            "Address": "; ".join(address_parts),
            "Phone": phone_str,
            "Email / Website": email_web,
        })
        i += 1

    return records


# --------------------------------------------------------------------------
# 6. Exporters
# --------------------------------------------------------------------------
def build_word_document(records) -> bytes:
    doc = Document()

    title = doc.add_heading("Extracted PDF Directory", level=1)
    title.style.font.color.rgb = RGBColor(0, 51, 102)

    current_category = None
    current_page = None
    for rec in records:
        if rec["Page"] != current_page:
            current_page = rec["Page"]
            p = doc.add_paragraph()
            run = p.add_run(f"— Page {current_page} —")
            run.italic = True
            run.font.size = Pt(9)
            run.font.color.rgb = RGBColor(112, 128, 144)

        if rec["Category"] != current_category:
            current_category = rec["Category"]
            h = doc.add_heading(current_category, level=2)
            h.style.font.color.rgb = RGBColor(180, 50, 50)

        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)

        name_run = p.add_run(rec["Name"] or "(unnamed entry)")
        name_run.bold = True
        name_run.font.size = Pt(11)
        name_run.font.name = "Calibri"

        details = []
        if rec["Address"]:
            details.append(rec["Address"])
        if rec["Phone"]:
            details.append(f"Tel: {rec['Phone']}")
        if rec["Email / Website"]:
            details.append(rec["Email / Website"])
        if details:
            detail_run = p.add_run("\n" + "  |  ".join(details))
            detail_run.font.size = Pt(10)
            detail_run.font.name = "Calibri"
            detail_run.font.color.rgb = RGBColor(60, 60, 60)

    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


def build_excel(records) -> bytes:
    df = pd.DataFrame(records, columns=["Page", "Category", "Name", "Address", "Phone", "Email / Website"])
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Directory")
        ws = writer.sheets["Directory"]
        widths = {"A": 7, "B": 22, "C": 32, "D": 40, "E": 20, "F": 28}
        for col, w in widths.items():
            ws.column_dimensions[col].width = w
    return buffer.getvalue()


# --------------------------------------------------------------------------
# 7. Streamlit UI
# --------------------------------------------------------------------------
def main():
    st.title("⚡ PDF Extraction Suite Pro")
    st.write(
        "Upload a scanned PDF directory. The app OCRs it, cleans the text, "
        "and lets you download a structured Word document and/or Excel "
        "spreadsheet."
    )

    with st.sidebar:
        st.header("Local tool paths")
        st.caption("Only needed on Windows if these tools aren't already on your PATH.")
        tesseract_cmd = st.text_input("Tesseract executable", value=DEFAULT_TESSERACT_CMD)
        poppler_path = st.text_input("Poppler `bin` folder (leave blank on macOS/Linux if installed via package manager)", value=DEFAULT_POPPLER_PATH)

        st.header("Speed / quality")
        dpi = st.slider("Render DPI", min_value=150, max_value=400, value=200, step=25,
                         help="200 is plenty for most printed directories. Only go to 300+ if text is genuinely tiny.")
        quality_mode = st.selectbox(
            "Preprocessing quality", ["Fast", "Balanced", "Max accuracy (slow)"], index=1,
            help="Fast = adaptive threshold only, best for clean printed pages. "
                 "Max accuracy adds a bilateral-filter denoise pass — only worth it on noisy photocopies.",
        )
        mode_map = {"Fast": "fast", "Balanced": "balanced", "Max accuracy (slow)": "max"}
        preprocess_mode = mode_map[quality_mode]

        cpu = os.cpu_count() or 4
        max_workers = st.slider(
            "Parallel OCR workers", min_value=1, max_value=cpu * 2, value=min(cpu, 8),
            help="OCR shells out to the Tesseract binary, so this can safely exceed your CPU core count.",
        )
        chunk_size = st.slider("Pages rendered per batch", min_value=5, max_value=50, value=15,
                                help="Smaller batches start OCR sooner and use less memory on large PDFs; larger batches render slightly faster overall.")
        page_limit = st.number_input(
            "Only process the first N pages (0 = all)", min_value=0, value=0, step=10,
            help="Handy for testing your settings on a big file before committing to the full run.",
        )

    export_choice = st.radio(
        "What do you want to download?",
        ["Word (.docx)", "Excel (.xlsx)", "Both"],
        horizontal=True,
    )

    uploaded_file = st.file_uploader("Upload scanned PDF", type=["pdf"])

    if uploaded_file is None:
        return

    st.info(f"Loaded file: **{uploaded_file.name}** ({uploaded_file.size / 1_000_000:.1f} MB)")

    if not st.button("Start processing", type="primary"):
        return

    if not os.path.isfile(tesseract_cmd):
        st.error(f"Tesseract not found at:\n`{tesseract_cmd}`\nFix the path in the sidebar and try again.")
        return

    pdf_bytes = uploaded_file.read()
    poppler_kwargs = {"poppler_path": poppler_path} if poppler_path.strip() else {}

    try:
        info = pdfinfo_from_bytes(pdf_bytes, **poppler_kwargs)
        total_pages = info["Pages"]
    except Exception as exc:
        st.error(f"Could not read the PDF. Check the Poppler path in the sidebar.\n\n{exc}")
        return

    pages_to_do = min(total_pages, page_limit) if page_limit else total_pages
    st.success(f"{total_pages} page(s) detected — processing {pages_to_do}.")

    progress_bar = st.progress(0.0)
    status_text = st.empty()
    start_time = time.time()

    results_by_page = {}
    completed = 0

    # Pipeline: render one batch of pages, immediately hand it to the OCR
    # thread pool, then render the NEXT batch while OCR is still chewing
    # through the previous one. This keeps peak memory to one batch's worth
    # of images (instead of the whole PDF) and overlaps rendering with OCR
    # instead of doing the two strictly back-to-back. Results are keyed by
    # page number and re-sorted afterwards — pages don't necessarily finish
    # OCR in submission order, but the parser below needs reading order.
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        pending_futures = {}

        for batch_start in range(1, pages_to_do + 1, chunk_size):
            batch_end = min(batch_start + chunk_size - 1, pages_to_do)
            status_text.text(f"Rendering pages {batch_start}\u2013{batch_end} of {pages_to_do}...")

            try:
                batch_images = convert_from_bytes(
                    pdf_bytes, dpi=dpi,
                    first_page=batch_start, last_page=batch_end,
                    grayscale=True,
                    thread_count=min(max_workers, chunk_size),
                    **poppler_kwargs,
                )
            except Exception as exc:
                st.error(f"Rendering failed on pages {batch_start}\u2013{batch_end}: {exc}")
                return

            for offset, img in enumerate(batch_images):
                page_num = batch_start + offset
                fut = executor.submit(ocr_single_page, (page_num, img), tesseract_cmd, preprocess_mode)
                pending_futures[fut] = page_num

            status_text.text(f"OCR running on pages {batch_start}\u2013{batch_end} (of {pages_to_do})...")

            # Drain any OCR results that have already finished so memory
            # from completed batches doesn't pile up.
            done_now = [f for f in list(pending_futures) if f.done()]
            for fut in done_now:
                page_num = pending_futures.pop(fut)
                results_by_page[page_num] = fut.result()
                completed += 1
                progress_bar.progress(completed / pages_to_do)

        # Wait for whatever OCR work is still outstanding.
        for fut in as_completed(pending_futures):
            page_num = pending_futures[fut]
            results_by_page[page_num] = fut.result()
            completed += 1
            progress_bar.progress(completed / pages_to_do)

    elapsed = time.time() - start_time
    status_text.text(
        f"OCR complete: {pages_to_do} pages in {elapsed:.0f}s "
        f"({pages_to_do / elapsed:.1f} pages/sec)."
    )

    # Reassemble in page order — OCR completion order has no relationship
    # to reading order, but the parser below depends on reading order.
    raw_lines = []
    for page_num in sorted(results_by_page):
        raw_lines.extend(results_by_page[page_num])

    records = parse_lines_to_records(raw_lines)

    if not records:
        st.warning("No text could be confidently extracted from this PDF. Try a higher DPI or check image quality.")
        return

    st.success(f"Extracted {len(records)} structured entries across {pages_to_do} pages.")
    df_preview = pd.DataFrame(records)
    st.subheader("Preview")
    st.dataframe(df_preview.head(25), use_container_width=True)

    base_filename = os.path.splitext(uploaded_file.name)[0]

    col1, col2 = st.columns(2)

    if export_choice in ("Word (.docx)", "Both"):
        docx_bytes = build_word_document(records)
        with col1:
            st.download_button(
                label="📥 Download Word (.docx)",
                data=docx_bytes,
                file_name=f"{base_filename}_EXTRACTED.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

    if export_choice in ("Excel (.xlsx)", "Both"):
        xlsx_bytes = build_excel(records)
        with col2:
            st.download_button(
                label="📥 Download Excel (.xlsx)",
                data=xlsx_bytes,
                file_name=f"{base_filename}_EXTRACTED.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    st.balloons()


if __name__ == "__main__":
    main()
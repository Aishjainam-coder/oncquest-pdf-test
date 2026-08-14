import sys
from pathlib import Path
sys.path.append(str(Path(__file__).resolve().parent.parent))

import fitz # PyMuPDF
from pdf2docx import Converter as PDF2DocxConverter
from converter import render_html_to_pdf_and_preview

html_path = Path("output/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.html")
pdf_path = Path("output/test_debug.pdf")
docx_path = Path("output/test_debug.docx")

# Step 1: Render HTML to PDF
print("Rendering HTML to PDF...")
render_html_to_pdf_and_preview(html_path, pdf_path)
print(f"PDF exists: {pdf_path.exists()}, Size: {pdf_path.stat().st_size if pdf_path.exists() else 0} bytes")

# Step 2: Extract text from PDF using PyMuPDF to see if there is text
if pdf_path.exists():
    doc = fitz.open(str(pdf_path))
    print(f"Pages in PDF: {len(doc)}")
    full_text = ""
    for idx, page in enumerate(doc):
        t = page.get_text()
        print(f"Page {idx+1} text length: {len(t)}")
        if t.strip():
            full_text += t + "\n"
    print("Total extracted characters:", len(full_text))
    doc.close()

# Step 3: Convert using pdf2docx
print("Converting using pdf2docx...")
try:
    cv = PDF2DocxConverter(str(pdf_path))
    cv.convert(str(docx_path))
    cv.close()
    print("pdf2docx conversion finished.")
except Exception as e:
    print("pdf2docx conversion failed:", e)

# Step 4: Check docx contents
if docx_path.exists():
    print(f"Docx file size: {docx_path.stat().st_size} bytes")

import fitz
from pathlib import Path

# Paths
input_pdf_path = Path("outsourcing pdf/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.pdf")
output_dir = Path("output")
output_pdf_path = output_dir / "TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee_compiled.pdf"

# Render input PDF
if input_pdf_path.exists():
    doc_in = fitz.open(str(input_pdf_path))
    for page_idx in range(min(3, len(doc_in))):
        page = doc_in[page_idx]
        pix = page.get_pixmap(dpi=150)
        pix.save(f"output/input_page_{page_idx+1}.png")
    doc_in.close()
    print("Rendered input PDF pages.")

# Render output PDF (if it exists)
# Wait, let's look for any compiled.pdf or generated PDF in output directory first
pdf_files = list(output_dir.glob("*.pdf"))
print("PDF files in output:", [p.name for p in pdf_files])

for pdf_p in pdf_files:
    doc_out = fitz.open(str(pdf_p))
    for page_idx in range(min(3, len(doc_out))):
        page = doc_out[page_idx]
        pix = page.get_pixmap(dpi=150)
        pix.save(f"output/output_{pdf_p.stem}_page_{page_idx+1}.png")
    doc_out.close()
    print(f"Rendered {pdf_p.name} pages.")

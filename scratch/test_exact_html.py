import sys
from pathlib import Path
sys.path.append(str(Path(__file__).resolve().parent.parent))

import fitz
from converter import render_exact_pdf_layout_html, render_html_to_pdf_and_preview

pdf_path = Path("outsourcing pdf/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.pdf")
html_out = Path("output/exact_position_test.html")
pdf_out = Path("output/exact_position_test_compiled.pdf")

# Generate HTML using render_exact_pdf_layout_html
doc = fitz.open(str(pdf_path))
html_content = render_exact_pdf_layout_html(doc, doc_title=pdf_path.name)
html_out.write_text(html_content, encoding="utf-8")
doc.close()
print("Generated exact layout HTML.")

# Compile HTML to PDF
render_html_to_pdf_and_preview(html_out, pdf_out)
print(f"Compiled PDF exists: {pdf_out.exists()}, Size: {pdf_out.stat().st_size if pdf_out.exists() else 0} bytes")

# Inspect drawings of compiled PDF
if pdf_out.exists():
    doc_compiled = fitz.open(str(pdf_out))
    page = doc_compiled[0]
    drawings = page.get_drawings()
    print(f"Total drawings in compiled PDF (Page 1): {len(drawings)}")
    rect_drawings = [d for d in drawings if d.get('rect')]
    print(f"Drawings with rect: {len(rect_drawings)}")
    for i, d in enumerate(rect_drawings[:15]):
        rect = d['rect']
        print(f"  {i+1}: rect=({rect.x0:.1f}, {rect.y0:.1f}, {rect.x1:.1f}, {rect.y1:.1f}) w={rect.width:.1f} h={rect.height:.1f} fill={d.get('fill')} color={d.get('color')} stroke_w={d.get('width')}")
    doc_compiled.close()

import fitz
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')
input_pdf = Path("outsourcing pdf/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.pdf")
output_pdf = Path("output/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee_compiled.pdf")

def inspect_page(pdf_path, name):
    doc = fitz.open(str(pdf_path))
    page = doc[0]
    drawings = page.get_drawings()
    print(f"\n===== {name} (Page 1) =====")
    print(f"Total drawings: {len(drawings)}")
    rect_drawings = [d for d in drawings if d.get('rect')]
    print(f"Drawings with rect: {len(rect_drawings)}")
    for i, d in enumerate(rect_drawings[:15]):
        rect = d['rect']
        print(f"  {i+1}: rect=({rect.x0:.1f}, {rect.y0:.1f}, {rect.x1:.1f}, {rect.y1:.1f}) w={rect.width:.1f} h={rect.height:.1f} fill={d.get('fill')} color={d.get('color')} stroke_w={d.get('width')}")
    doc.close()

inspect_page(input_pdf, "Input PDF")
inspect_page(output_pdf, "Output Compiled PDF")

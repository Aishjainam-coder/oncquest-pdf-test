import fitz
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')
pdf_path = Path("outsourcing pdf/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.pdf")
doc = fitz.open(str(pdf_path))

print(f"Total pages: {len(doc)}")
for idx in range(min(5, len(doc))):
    page = doc[idx]
    drawings = page.get_drawings()
    print(f"\n--- Page {idx+1} (Total drawings: {len(drawings)}) ---")
    rect_drawings = [d for d in drawings if d.get('rect')]
    print(f"Drawings with rect: {len(rect_drawings)}")
    
    # Print the first 10 drawings' properties
    for d_idx, d in enumerate(rect_drawings[:15]):
        rect = d['rect']
        print(f"  Drawing {d_idx+1}: rect=({rect.x0:.2f}, {rect.y0:.2f}, {rect.x1:.2f}, {rect.y1:.2f}) "
              f"width={rect.width:.2f} height={rect.height:.2f} fill={d.get('fill')} color={d.get('color')} width_stroke={d.get('width')}")

import fitz
from pathlib import Path

pdf_path = Path("outsourcing pdf/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.pdf")
doc = fitz.open(str(pdf_path))
page = doc[0]
drawings = page.get_drawings()

print("Sample drawing dictionary keys and values:")
for idx, d in enumerate(drawings[:5]):
    print(f"\nDrawing {idx+1}:")
    for k, v in d.items():
        if k == 'items':
            print(f"  {k}: length={len(v)}, first={v[0] if v else None}")
        else:
            print(f"  {k}: {v}")
doc.close()

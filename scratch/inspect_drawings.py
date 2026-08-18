import fitz
from pathlib import Path

pdf_path = Path(r"c:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main\outsourcing pdf\TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.pdf")

doc = fitz.open(pdf_path)
page = doc[0]

print("Drawings on Page 1:")
drawings = page.get_drawings()
print(f"Total drawings: {len(drawings)}")

# Filter drawings around the test name vertical range (y: 200 to 245)
for i, d in enumerate(drawings):
    bbox = d["rect"]
    # Check if bbox overlaps with the vertical region of the test name
    if 200 <= bbox[1] <= 245:
        print(f"Drawing {i}: rect={bbox} type={d['type']} items={len(d['items'])}")
        for item in d["items"]:
            print(f"  Item: {item}")

doc.close()

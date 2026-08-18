import fitz
from pathlib import Path

pdf_path = Path(r"c:\Users\aishwarya.jain\Downloads\TestReport_SAROJ DEVI (OQG2604250052)_2600133939_1eb3f888-f97e-4181-b84d-d9536de7af26_output.pdf")

doc = fitz.open(pdf_path)
print(f"Saroj Devi PDF has {len(doc)} pages.")

# Search for any page that has patient table or test name keywords
for i in range(len(doc)):
    page_text = doc[i].get_text()
    if "test name" in page_text.lower() or "patient details" in page_text.lower() or "sex" in page_text.lower() or "age" in page_text.lower():
        print(f"Match on Page {i+1}:")
        for b in doc[i].get_text("blocks"):
            x0, y0, x1, y1, text, block_no, block_type = b
            if "test name" in text.lower() or "patient" in text.lower() or "sex" in text.lower() or "age" in text.lower():
                print(f"  [{x0:.1f}, {y0:.1f}, {x1:.1f}, {y1:.1f}] {text.strip().replace('\n', ' ')}")

doc.close()

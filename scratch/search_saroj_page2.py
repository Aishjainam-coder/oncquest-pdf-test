import fitz
import sys
from pathlib import Path

# Ensure UTF-8 output
sys.stdout.reconfigure(encoding='utf-8')

pdf_path = Path(r"c:\Users\aishwarya.jain\Downloads\TestReport_SAROJ DEVI (OQG2604250052)_2600133939_1eb3f888-f97e-4181-b84d-d9536de7af26_output.pdf")

doc = fitz.open(pdf_path)
page = doc[1]  # Page 2 (0-indexed is page 1, 1-indexed is page 2)
print("Page 2 Height:", page.rect.height)
print("All text blocks on Page 2:")
for b in page.get_text("blocks"):
    x0, y0, x1, y1, text, block_no, block_type = b
    text_clean = text.strip().replace("\n", " ")
    if text_clean:
        print(f"  [{x0:.1f}, {y0:.1f}, {x1:.1f}, {y1:.1f}] {text_clean}")
doc.close()

import fitz
import sys
from pathlib import Path

pdf_paths = [
    Path(r"c:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main\outsourcing pdf\TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.pdf"),
    Path(r"c:\Users\aishwarya.jain\Downloads\TestReport_BASANTI DEVI 10078040 OQG2604170493_60400127778_42f516a3-ebf4-4187-a505-e2b8c4b5e928_target_output.pdf"),
    Path(r"c:\Users\aishwarya.jain\Downloads\TestReport_BIRENDRA KR TRIPATHI (KOL2604200324)_2600132470_e5d40a68-7e61-43d6-aae0-577a57bb598a_target_output.pdf"),
    Path(r"c:\Users\aishwarya.jain\Downloads\TestReport_DEVIKA 2908914 OQG2603250447_60300129567_39c65aa4-7a37-4af2-a4e6-46c303a638fe_target_output.pdf"),
    Path(r"c:\Users\aishwarya.jain\Downloads\TestReport_SAROJ DEVI (OQG2604250052)_2600133939_1eb3f888-f97e-4181-b84d-d9536de7af26_output.pdf"),
    Path(r"c:\Users\aishwarya.jain\Downloads\vendor pdf_output.pdf")
]

for p in pdf_paths:
    if not p.exists():
        print(f"Path does not exist: {p}")
        continue
    print("="*80)
    print(f"File: {p.name}")
    print("="*80)
    try:
        doc = fitz.open(p)
        print(f"Number of pages: {len(doc)}")
        page = doc[0]
        print("Page height:", page.rect.height)
        print("All text blocks on page 1:")
        for b in page.get_text("blocks"):
            x0, y0, x1, y1, text, block_no, block_type = b
            text_clean = text.strip().replace("\n", " ")
            if text_clean:
                print(f"  [{x0:.1f}, {y0:.1f}, {x1:.1f}, {y1:.1f}] {text_clean}")
        doc.close()
    except Exception as e:
        print(f"Error reading {p.name}: {e}")
    print("\n")

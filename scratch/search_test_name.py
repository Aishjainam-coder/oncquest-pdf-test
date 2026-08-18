import fitz
from pathlib import Path
import re

pdf_paths = [
    Path(r"c:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main\outsourcing pdf\TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.pdf"),
    Path(r"c:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main\output\TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee_compiled.pdf"),
    Path(r"c:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main\output\exact_position_test_compiled.pdf"),
    Path(r"c:\Users\aishwarya.jain\Downloads\TestReport_BASANTI DEVI 10078040 OQG2604170493_60400127778_42f516a3-ebf4-4187-a505-e2b8c4b5e928_target_output.pdf"),
    Path(r"c:\Users\aishwarya.jain\Downloads\TestReport_BIRENDRA KR TRIPATHI (KOL2604200324)_2600132470_e5d40a68-7e61-43d6-aae0-577a57bb598a_target_output.pdf"),
    Path(r"c:\Users\aishwarya.jain\Downloads\TestReport_DEVIKA 2908914 OQG2603250447_60300129567_39c65aa4-7a37-4af2-a4e6-46c303a638fe_target_output.pdf"),
    Path(r"c:\Users\aishwarya.jain\Downloads\TestReport_SAROJ DEVI (OQG2604250052)_2600133939_1eb3f888-f97e-4181-b84d-d9536de7af26_output.pdf"),
    Path(r"c:\Users\aishwarya.jain\Downloads\vendor pdf_output.pdf")
]

for p in pdf_paths:
    if not p.exists():
        continue
    print("="*80)
    print(f"File: {p.name}")
    print("="*80)
    doc = fitz.open(p)
    
    # Let's search for "Test Name" pattern or look at page 1 blocks
    page = doc[0]
    blocks = page.get_text("dict")["blocks"]
    
    for block_idx, block in enumerate(blocks):
        if "lines" not in block:
            continue
        for line_idx, line in enumerate(block["lines"]):
            for span_idx, span in enumerate(line["spans"]):
                text = span["text"]
                # Print any spans with known test name indicators or around typical coordinates
                if "Test Name" in text or "ORION" in text or "Liquidseq" in text or "HEREDITARY" in text or "PERFORMED." in text:
                    print(f"Block {block_idx}, Line {line_idx}, Span {span_idx}:")
                    print(f"  Text: '{text}'")
                    print(f"  Bbox: {span['bbox']}")
                    print(f"  Font: '{span['font']}' Size: {span['size']:.2f} Color: {span['color']} Flags: {span['flags']}")
    
    doc.close()
    print()

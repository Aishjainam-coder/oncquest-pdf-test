import fitz
from pathlib import Path
import sys
import re

# Ensure UTF-8 output
sys.stdout.reconfigure(encoding='utf-8')

workspace_dir = Path(r"c:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main")
downloads_dir = Path(r"c:\Users\aishwarya.jain\Downloads")

unique_pdfs = [
    workspace_dir / "output/exact_position_test_compiled.pdf",
    workspace_dir / "output/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee_compiled.pdf",
    workspace_dir / "outsourcing pdf/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.pdf",
    downloads_dir / "TestReport_BASANTI DEVI 10078040 OQG2604170493_60400127778_42f516a3-ebf4-4187-a505-e2b8c4b5e928_target_output.pdf",
    downloads_dir / "TestReport_BIRENDRA KR TRIPATHI (KOL2604200324)_2600132470_e5d40a68-7e61-43d6-aae0-577a57bb598a_target_output.pdf",
    downloads_dir / "TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.pdf",
    downloads_dir / "vendor pdf_output (1).pdf",
    downloads_dir / "vendor pdf_output (2).pdf",
    downloads_dir / "vendor pdf_output (3).pdf",
    downloads_dir / "vendor pdf_output.pdf"
]

print("="*80)
print("VERIFYING REPLACEMENTS IN ALL MODIFIED PDF FILES")
print("="*80)

for p in unique_pdfs:
    if not p.exists():
        print(f"File not found: {p}")
        continue
    
    try:
        doc = fitz.open(p)
    except Exception as e:
        print(f"Error opening {p.name}: {e}")
        continue
        
    print(f"\nFile: {p.name}")
    page = doc[0]
    blocks = page.get_text("dict")["blocks"]
    
    found_replacement = False
    for b_idx, b in enumerate(blocks):
        if "lines" not in b:
            continue
        for l_idx, l in enumerate(b["lines"]):
            for s_idx, s in enumerate(l["spans"]):
                text = s["text"]
                # Normalize whitespace (replaces \xa0 with space)
                norm_text = re.sub(r'\s+', ' ', text)
                
                if "TEST NAME" in norm_text:
                    print(f"  Page 1, Block {b_idx}, Line {l_idx}, Span {s_idx}:")
                    print(f"    Text: '{text}' (normalized: '{norm_text}')")
                    print(f"    Bbox: {s['bbox']}")
                    print(f"    Font: '{s['font']}' Size: {s['size']:.2f} Color: {s['color']}")
                    found_replacement = True
                    
                if any(kw in text for kw in ["Liquidseq Actionable", "Illumina Novaseq", "HEREDITARY CANCER", "ORION (V2)"]):
                    print(f"  [ERROR] Found old test name text: '{text}'")
                    
    if not found_replacement:
        print("  [WARNING] 'TEST NAME' placeholder not found on page 1!")
    doc.close()

print("\nVerification completed!")

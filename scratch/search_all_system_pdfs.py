import fitz
from pathlib import Path
import sys

# Ensure UTF-8 output
sys.stdout.reconfigure(encoding='utf-8')

workspace_dir = Path(r"c:\Users\aishwarya.jain\Downloads\Oncquest-main")
downloads_dir = Path(r"c:\Users\aishwarya.jain\Downloads")

all_pdfs = list(workspace_dir.glob("**/*.pdf")) + list(downloads_dir.glob("*.pdf"))
# Remove duplicates
unique_pdfs = []
seen = set()
for p in all_pdfs:
    res_path = p.resolve()
    if res_path not in seen:
        seen.add(res_path)
        unique_pdfs.append(p)

print(f"Found {len(unique_pdfs)} unique PDF files in workspace & downloads.")

for p in unique_pdfs:
    if not p.exists():
        continue
    
    try:
        doc = fitz.open(p)
    except Exception as e:
        print(f"Error opening {p.name}: {e}")
        continue
        
    print("-" * 80)
    print(f"File: {p}")
    print(f"Pages: {len(doc)}")
    
    found_any = False
    for page_idx in range(len(doc)):
        page = doc[page_idx]
        blocks = page.get_text("dict")["blocks"]
        for block_idx, block in enumerate(blocks):
            if "lines" not in block:
                continue
            for line_idx, line in enumerate(block["lines"]):
                for span_idx, span in enumerate(line["spans"]):
                    text = span["text"]
                    # Look for test name indicators
                    if "Test Name" in text or "ORION" in text or "Liquidseq" in text or "HEREDITARY" in text or "PERFORMED." in text:
                        print(f"  Page {page_idx+1}, Block {block_idx}, Line {line_idx}, Span {span_idx}:")
                        print(f"    Text: '{text}'")
                        print(f"    Bbox: {span['bbox']}")
                        print(f"    Font: '{span['font']}' Size: {span['size']:.2f} Color: {span['color']} Flags: {span['flags']}")
                        found_any = True
    if not found_any:
        print("  No test name patterns matched.")
    doc.close()
    print()

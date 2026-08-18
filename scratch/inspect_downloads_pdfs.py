import fitz
import sys
from pathlib import Path

downloads_dir = Path(r"c:\Users\aishwarya.jain\Downloads")
pdf_files = list(downloads_dir.glob("*.pdf"))

workspace_pdf_files = [
    Path(r"c:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main\outsourcing pdf\TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.pdf"),
    Path(r"c:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main\output\TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee_compiled.pdf"),
    Path(r"c:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main\output\exact_position_test_compiled.pdf"),
    Path(r"c:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main\output\test_debug.pdf")
]

all_pdfs = workspace_pdf_files + pdf_files

output_file = Path(r"c:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main\scratch\downloads_inspection.txt")

with open(output_file, "w", encoding="utf-8") as f:
    for p in all_pdfs:
        if not p.exists():
            f.write(f"Path does not exist: {p}\n\n")
            continue
        f.write("="*80 + "\n")
        f.write(f"File: {p}\n")
        f.write("="*80 + "\n")
        try:
            doc = fitz.open(p)
            f.write(f"Number of pages: {len(doc)}\n")
            page = doc[0]
            f.write(f"Page height: {page.rect.height}\n")
            f.write("All text blocks on page 1:\n")
            for b in page.get_text("blocks"):
                x0, y0, x1, y1, text, block_no, block_type = b
                text_clean = text.strip().replace("\n", " ")
                if text_clean:
                    f.write(f"  [{x0:.1f}, {y0:.1f}, {x1:.1f}, {y1:.1f}] {text_clean}\n")
            doc.close()
        except Exception as e:
            f.write(f"Error reading {p.name}: {e}\n")
        f.write("\n\n")

print(f"Inspection written to {output_file}")

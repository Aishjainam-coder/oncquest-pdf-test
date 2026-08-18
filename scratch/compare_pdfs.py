import fitz
import sys
from pathlib import Path
import difflib

def extract_pdf_text_by_page(pdf_path):
    doc = fitz.open(pdf_path)
    pages_text = []
    for i, page in enumerate(doc):
        # Extract text blocks or plain text
        text = page.get_text("text")
        pages_text.append(text)
    doc.close()
    return pages_text

def compare_pdfs(original_path, compiled_path, output_report_path):
    orig_pages = extract_pdf_text_by_page(original_path)
    comp_pages = extract_pdf_text_by_page(compiled_path)
    
    with open(output_report_path, "w", encoding="utf-8") as f:
        f.write(f"Original PDF: {original_path}\n")
        f.write(f"Compiled PDF: {compiled_path}\n")
        f.write(f"Original pages: {len(orig_pages)}\n")
        f.write(f"Compiled pages: {len(comp_pages)}\n\n")
        
        max_pages = max(len(orig_pages), len(comp_pages))
        
        for page_idx in range(max_pages):
            page_num = page_idx + 1
            f.write(f"\n=========================================\n")
            f.write(f"--- Page {page_num} Comparison ---\n")
            f.write(f"=========================================\n")
            
            if page_idx >= len(orig_pages):
                f.write(f"[Page {page_num}] Original page missing. Compiled page content:\n")
                f.write(comp_pages[page_idx])
                f.write("\n")
                continue
                
            if page_idx >= len(comp_pages):
                f.write(f"[Page {page_num}] Compiled page missing. Original page content:\n")
                f.write(orig_pages[page_idx])
                f.write("\n")
                continue
                
            orig_text = orig_pages[page_idx]
            comp_text = comp_pages[page_idx]
            
            orig_lines = [line.strip() for line in orig_text.splitlines() if line.strip()]
            comp_lines = [line.strip() for line in comp_text.splitlines() if line.strip()]
            
            # Compare line-by-line using difflib
            diff = list(difflib.unified_diff(
                orig_lines,
                comp_lines,
                fromfile=f"Original_Page_{page_num}",
                tofile=f"Compiled_Page_{page_num}",
                lineterm=""
            ))
            
            if not diff:
                f.write(f"[Page {page_num}] Text matches exactly!\n")
            else:
                f.write(f"[Page {page_num}] Differences found:\n")
                for line in diff:
                    f.write(f"{line}\n")

if __name__ == "__main__":
    orig = r"c:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main\outsourcing pdf\TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.pdf"
    comp = r"c:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main\output\TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee_compiled.pdf"
    output_report = r"c:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main\scratch\compare_output.txt"
    compare_pdfs(orig, comp, output_report)
    print(f"Comparison report written to {output_report}")

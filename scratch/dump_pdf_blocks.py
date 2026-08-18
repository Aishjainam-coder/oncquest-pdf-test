import fitz
import sys

def dump_page_blocks(pdf_path, txt_output_path):
    doc = fitz.open(pdf_path)
    with open(txt_output_path, "w", encoding="utf-8") as f:
        f.write(f"PDF Path: {pdf_path}\n")
        f.write(f"Total Pages: {len(doc)}\n\n")
        
        for page_idx in range(len(doc)):
            page = doc[page_idx]
            f.write(f"=========================================\n")
            f.write(f"PAGE {page_idx + 1}\n")
            f.write(f"=========================================\n")
            
            # Extract text blocks
            blocks = page.get_text("blocks")
            # Sort blocks primarily by y0 (top), then x0 (left)
            blocks.sort(key=lambda b: (round(b[1], 1), round(b[0], 1)))
            
            f.write(f"Total Blocks: {len(blocks)}\n")
            for b_idx, block in enumerate(blocks):
                x0, y0, x1, y1, text, block_no, block_type = block
                clean_text = text.replace("\n", " ").strip()
                f.write(f"  Block {b_idx + 1} (bbox: [{x0:.1f}, {y0:.1f}, {x1:.1f}, {y1:.1f}], type: {block_type}):\n")
                f.write(f"    \"{clean_text}\"\n")
            f.write("\n")
    doc.close()

if __name__ == "__main__":
    orig_pdf = r"outsourcing pdf/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.pdf"
    comp_pdf = r"output/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee_compiled.pdf"
    
    dump_page_blocks(orig_pdf, r"scratch/orig_blocks.txt")
    dump_page_blocks(comp_pdf, r"scratch/comp_blocks.txt")
    print("Dumps generated in scratch/orig_blocks.txt and scratch/comp_blocks.txt")

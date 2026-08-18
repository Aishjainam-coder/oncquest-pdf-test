import fitz

pdf_path = r"c:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main\scratch\test_output.pdf"

doc = fitz.open(pdf_path)
page = doc[0]

print("All text blocks on Page 1 of test output:")
for b in page.get_text("blocks"):
    x0, y0, x1, y1, text, block_no, block_type = b
    text_clean = text.strip().replace("\n", " ")
    if text_clean:
        print(f"  [{x0:.1f}, {y0:.1f}, {x1:.1f}, {y1:.1f}] {text_clean}")
doc.close()

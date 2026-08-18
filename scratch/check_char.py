import fitz

pdf_path = r"c:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main\scratch\test_output.pdf"
doc = fitz.open(pdf_path)
page = doc[0]

for b in page.get_text("blocks"):
    text = b[4].strip()
    if "TEST" in text:
        print(f"Text block: {repr(text)}")
        for char in text:
            print(f"  Char: {repr(char)} Unicode: {ord(char)}")
doc.close()

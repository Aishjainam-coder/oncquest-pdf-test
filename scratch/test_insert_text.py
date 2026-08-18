import fitz

doc = fitz.open()
page = doc.new_page()
rect = fitz.Rect(100, 100, 300, 120)

# Try insert_textbox
page.insert_textbox(rect, "TEST NAME BOX", fontname="helv", fontsize=12)

# Try insert_text
point = fitz.Point(100, 150)
page.insert_text(point, "TEST NAME TEXT", fontname="helv", fontsize=12)

# Save and check the text blocks
doc.save("scratch/test_spaces.pdf")
doc.close()

doc2 = fitz.open("scratch/test_spaces.pdf")
page2 = doc2[0]
for b in page2.get_text("blocks"):
    text = b[4].strip()
    print(f"Text block: {repr(text)}")
    for char in text:
        print(f"  Char: {repr(char)} (Unicode: {ord(char)})")
doc2.close()

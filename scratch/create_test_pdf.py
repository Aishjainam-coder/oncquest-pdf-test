import fitz

input_pdf = "backup_pdfs/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.pdf"
output_pdf = "scratch/sng_test_input.pdf"

print(f"Opening base PDF: {input_pdf}")
doc = fitz.open(input_pdf)
page = doc[0]

# Add test text in a visible area (e.g. at position x=100, y=250)
rect = fitz.Rect(100, 250, 500, 300)
# We draw a white rectangle to clear the background first
page.draw_rect(rect, color=(1, 1, 1), fill=(1, 1, 1))

# Insert SNG Gen's Lab pvt ltd text
text = "This report is issued by SNG Gen's Lab pvt ltd. and verified."
print(f"Inserting text: '{text}'")
page.insert_textbox(rect, text, fontsize=11, fontname="helv", color=(0, 0, 0))

# Also insert a table header or key-value pair text to test different formats
rect2 = fitz.Rect(100, 300, 500, 350)
page.draw_rect(rect2, color=(1, 1, 1), fill=(1, 1, 1))
text2 = "Laboratory: SNG Gen's Lab pvt ltd"
page.insert_textbox(rect2, text2, fontsize=11, fontname="helv", color=(0, 0, 0))

doc.save(output_pdf)
doc.close()
print(f"Successfully saved test PDF to: {output_pdf}")

import fitz

doc = fitz.open('outsourcing pdf/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.pdf')
page = doc[0]
H = page.rect.height

print("Page height:", H)
print("All text blocks on page 1:")
for b in page.get_text("blocks"):
    x0, y0, x1, y1, text, block_no, block_type = b
    text_clean = text.strip().replace("\n", " ")
    if text_clean:
        print(f"[{x0:.1f}, {y0:.1f}, {x1:.1f}, {y1:.1f}] {text_clean}")

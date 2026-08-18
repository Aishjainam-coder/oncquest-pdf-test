import fitz

color_int = 32896
# Convert color_int to RGB tuple of float
# Let's inspect how PyMuPDF extracts RGB from this integer.
# If color_int is an sRGB integer, it could be (R, G, B) = ((color >> 16) & 255, (color >> 8) & 255, color & 255)
# Let's check with some standard conversion.
r = (color_int >> 16) & 255
g = (color_int >> 8) & 255
b = color_int & 255
print(f"Option 1 (RGB): {r, g, b}")

# In PDF, colors are often stored as BGR or RGB.
# Let's see if fitz has a helper.
print("sRGB to RGB float:")
# Let's inspect a PDF where color is 32896, and see what color is in the PDF.
doc = fitz.open(r"c:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main\outsourcing pdf\TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.pdf")
page = doc[0]
blocks = page.get_text("dict")["blocks"]
for b in blocks:
    if "lines" in b:
        for l in b["lines"]:
            for s in l["spans"]:
                if "Liquidseq" in s["text"]:
                    print(f"Found span '{s['text']}' with color={s['color']}")
                    # Let's convert color to RGB components using PyMuPDF's built-in conversion:
                    # In PyMuPDF, span color is an integer. Let's see how it is converted to RGB.
                    # fitz.utils.getColor or tuple representation
                    # RGB color tuple:
                    rgb_tuple = [((s['color'] >> 16) & 255) / 255.0, ((s['color'] >> 8) & 255) / 255.0, (s['color'] & 255) / 255.0]
                    print(f"Converted float RGB: {rgb_tuple}")
doc.close()

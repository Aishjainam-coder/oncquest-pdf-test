import fitz
from pathlib import Path

pdf_path = Path(r"c:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main\outsourcing pdf\TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.pdf")
output_path = Path(r"c:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main\scratch\test_output.pdf")

doc = fitz.open(pdf_path)
page = doc[0]

blocks = page.get_text("dict")["blocks"]
rect1, rect2 = None, None
font_size = 14.0
color_int = 32896

for b in blocks:
    if "lines" not in b:
        continue
    for l in b["lines"]:
        for s in l["spans"]:
            if "Liquidseq Actionable Genomic Profiling Panel" in s["text"]:
                rect1 = fitz.Rect(s["bbox"])
                font_size = s["size"]
                color_int = s["color"]
            elif "On Illumina Novaseq 6000 Platform" in s["text"]:
                rect2 = fitz.Rect(s["bbox"])

if rect1 and rect2:
    print(f"Found spans. Rect1: {rect1}, Rect2: {rect2}")
    
    # Redact original text
    page.add_redact_annot(rect1)
    page.add_redact_annot(rect2)
    page.apply_redactions(graphics=0)
    
    r = ((color_int >> 16) & 255) / 255.0
    g = ((color_int >> 8) & 255) / 255.0
    b = (color_int & 255) / 255.0
    color_rgb = (r, g, b)
    
    font_path = "C:/Windows/Fonts/cambriaz.ttf"
    font_name = "Cambria-BoldItalic"
    
    # Expand the target rect vertically to avoid layout overflow
    rect1.y0 -= 3
    rect1.y1 += 3
    
    ret = page.insert_textbox(
        rect1,
        "TEST NAME",
        fontname=font_name,
        fontfile=font_path,
        fontsize=font_size,
        color=color_rgb,
        align=1
    )
    print(f"insert_textbox returned: {ret}")
    
    # Save output
    doc.save(output_path)
    print("Saved test output.")
else:
    print("Error: spans not found.")
doc.close()

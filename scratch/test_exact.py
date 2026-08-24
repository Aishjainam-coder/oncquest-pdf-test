import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

import pymupdf as fitz
import json
from converter import render_exact_pdf_layout_html

# Load theme
theme_config = {}
theme_json_path = Path("theme.json")
if theme_json_path.exists():
    with open(theme_json_path, "r", encoding="utf-8") as f:
        theme_config = json.load(f)

# Open PDF
pdf_path = Path("outsourcing pdf/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.pdf")
doc = fitz.open(pdf_path)

# Render HTML
html_content = render_exact_pdf_layout_html(doc, doc_title=pdf_path.name, theme_config=theme_config)
doc.close()

# Save output
output_path = Path("output/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee_exact_test.html")
output_path.write_text(html_content, encoding="utf-8")
print(f"Successfully generated: {output_path}")

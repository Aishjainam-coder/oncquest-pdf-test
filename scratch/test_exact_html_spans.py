import json
import re
import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from converter import generate_dynamic_template_html, render_exact_pdf_layout_html
import pymupdf as fitz

with open('extracted_jsons/TestReport_SAROJ DEVI (OQG2604250052)_2600133939_1eb3f888-f97e-4181-b84d-d9536de7af26.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

html = generate_dynamic_template_html(data)

# Find all yellow highlights in HTML
matches = re.findall(r'<[^>]*background-color:\s*(?:#ffff00|yellow)[^>]*>(.*?)</[^>]+>', html, re.IGNORECASE)
print(f"Total HTML yellow highlights found: {len(matches)}")
for m in matches:
    print(f"  Highlighted content: '{m}'")

pdf_path = 'outsourcing pdf/TestReport_SAROJ DEVI (OQG2604250052)_2600133939_1eb3f888-f97e-4181-b84d-d9536de7af26.pdf'
if os.path.exists(pdf_path):
    with fitz.open(pdf_path) as doc_pdf:
        html_exact = render_exact_pdf_layout_html(doc_pdf)
    matches_exact = re.findall(r'<[^>]*background-color:\s*(?:#ffff00|yellow)[^>]*>(.*?)</[^>]+>', html_exact, re.IGNORECASE)
    print(f"\nTotal Exact PDF HTML yellow highlights found: {len(matches_exact)}")
    for m in matches_exact:
        print(f"  Highlighted exact content: '{m}'")

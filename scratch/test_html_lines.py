import sys
sys.path.insert(0, '.')
import fitz, re
from converter import render_exact_pdf_layout_html

doc = fitz.open('outsourcing pdf/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.pdf')
html = render_exact_pdf_layout_html(doc)
p1 = html[:html.find("id='page-2'")]

print("=== All HTML elements between 195pt and 255pt on Page 1 ===")
for line in p1.splitlines():
    m = re.search(r'top:([0-9.]+)pt', line)
    if m:
        top_val = float(m.group(1))
        if 195 <= top_val <= 255:
            print(f"{top_val:5.1f} pt: {line.strip()[:110]}")

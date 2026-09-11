import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import json
import re
from converter import generate_dynamic_template_html

with open('extracted_jsons/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

html = generate_dynamic_template_html(data)
matches = re.findall(r'<[^>]*background(?:-color)?:\s*(?:#ffff00|yellow)[^>]*>(.*?)</[^>]+>', html, re.IGNORECASE)
print(f"Total HTML yellow highlights found: {len(matches)}")
for m in matches:
    print(f"  Highlighted content: '{m}'")

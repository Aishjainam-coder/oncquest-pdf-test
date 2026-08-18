import json
import fitz

pdf_path = r"outsourcing pdf/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.pdf"
json_path = r"extracted_jsons/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.json"

with open(json_path, "r", encoding="utf-8") as f:
    data = json.load(f)

print("JSON Keys:", data.keys())
if "pages" in data:
    print(f"Total pages in JSON: {len(data['pages'])}")
    for i, p in enumerate(data['pages']):
        print(f"Page {i+1} keys:", p.keys())
        print(f"  Tables count: {len(p.get('tables', []))}")
        print(f"  KV count: {len(p.get('key_values', {}))}")
        print(f"  Content boxes: {len(p.get('content_boxes', []))}")
        print(f"  Text blocks count: {len(p.get('text_blocks', []))}")

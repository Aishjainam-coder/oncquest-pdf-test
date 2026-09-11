import json

with open('extracted_jsons/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.json', 'r', encoding='utf-8', errors='ignore') as f:
    data = json.load(f)

pages = data.get('pages') or data.get('document', {}).get('pages', [])
p0 = pages[0]
for idx, el in enumerate(p0.get('elements', [])):
    txt = str(el.get('text', '')) + str(el.get('data', '')) + str(el.get('title', ''))
    print(f"{idx:2d}: type={el.get('type', 'none'):12s} bbox={el.get('bbox')} text={repr(txt[:60])}")

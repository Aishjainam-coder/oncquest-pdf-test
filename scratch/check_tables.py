import json, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

d = json.load(open('extracted_jsons/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.json','r',encoding='utf-8'))

# Check tables with wrong headers
print("=== TABLE ANALYSIS ===")
for i, c in enumerate(d['content']):
    if c.get('type') == 'table':
        headers = c.get("headers",[])
        rows = c.get("rows",[])
        page = c.get("page")
        bbox = c.get("bbox",[])
        
        # Check if headers look like data (long text, first row of actual data)
        header_text = " ".join(str(h) for h in headers)
        print(f"\nTable {i} (page {page}):")
        print(f"  bbox: {bbox}")
        print(f"  headers ({len(headers)}): {headers}")
        print(f"  rows ({len(rows)})")
        if rows:
            print(f"  first row: {rows[0]}")
        
        # Flag if header seems wrong
        if any(len(str(h)) > 100 for h in headers):
            print("  >>> WARNING: Header seems like data (too long)!")
        if any(str(h).startswith(('X', 'GH-', 'olaparib')) for h in headers):
            print("  >>> WARNING: Header seems like data row!")

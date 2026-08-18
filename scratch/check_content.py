import json, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

d = json.load(open('extracted_jsons/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.json','r',encoding='utf-8'))

# Check all content items
for i, c in enumerate(d['content']):
    ctype = c.get("type")
    title = c.get("title","")[:80]
    page = c.get("page")
    print(f"Item {i}: type={ctype}, page={page}, title='{title}'")
    if ctype == 'table':
        headers = c.get("headers",[])
        rows = c.get("rows",[])
        print(f"  headers={headers}")
        print(f"  rows count={len(rows)}")
        if rows:
            print(f"  first_row={rows[0]}")
    elif ctype in ('box', 'banner'):
        ct = c.get("content_text",[])
        if isinstance(ct, list):
            print(f"  content_text lines={len(ct)}")
            if ct:
                print(f"  first line='{ct[0][:100]}'")
        else:
            print(f"  content_text='{str(ct)[:100]}'")
    print()

# Check pages structure
print("\n=== PAGES ===")
for p in d['pages']:
    pn = p.get("page_number")
    tables = p.get("tables",[])
    boxes = p.get("boxes_and_sections",[])
    print(f"Page {pn}: tables={len(tables)}, boxes={len(boxes)}")
    for t in tables:
        headers = t.get("headers",[])
        rows = t.get("rows",[])
        print(f"  Table: headers={headers}, rows={len(rows)}")
    for b in boxes:
        bt = b.get("type","")
        btitle = b.get("title","")[:60]
        print(f"  Box: type={bt}, title='{btitle}'")

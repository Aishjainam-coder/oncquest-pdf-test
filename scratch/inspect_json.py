import json

d = json.load(open(r'extracted_jsons/TestReport_AGATAMUDI VENKATA KAMANA (RJM2604120002)_2600130362_526f942c-1bc7-40ec-975b-0c41833042f4.json', 'r', encoding='utf-8'))
pages = d.get('pages', [])

for pi in range(min(5, len(pages))):
    p = pages[pi]
    boxes = p.get('boxes_and_sections', [])
    tables = p.get('tables', [])
    imgs = p.get('images_and_graphs', [])
    tbs = p.get('text_blocks', [])
    print(f"Page {pi+1}: boxes={len(boxes)}, tables={len(tables)}, images={len(imgs)}, text_blocks={len(tbs)}")
    for b in boxes:
        title = b.get("title", "")
        btype = b.get("type", "")
        print(f"  BOX: title={title!r} type={btype!r}")
    for t in tables:
        hdrs = t.get("headers", [])
        nrows = len(t.get("rows", []))
        print(f"  TABLE: headers={hdrs} rows={nrows}")

# Also check drawings on page 0
p0 = pages[0]
drawings = p0.get("drawings", [])
print(f"\nPage 1 drawings count: {len(drawings)}")
if drawings:
    print(f"First drawing keys: {list(drawings[0].keys())}")
    print(f"First drawing: {drawings[0]}")

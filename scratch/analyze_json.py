import json, os

fp = os.path.join("extracted_jsons", "TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.json")
d = json.load(open(fp, "r", encoding="utf-8"))

print("=== TOP-LEVEL KEYS ===")
for k in d:
    v = d[k]
    if isinstance(v, (list, dict, str)):
        print(f"  {k}: {type(v).__name__} (len={len(v)})")
    else:
        print(f"  {k}: {type(v).__name__} = {v}")

print("\n=== document.pages[0] element types ===")
for e in d["document"]["pages"][0]["elements"]:
    t = e["type"]
    extra = ""
    if t == "key_value":
        extra = f" keys={list(e.get('data',{}).keys())[:5]}"
    elif t == "table":
        cols = [c["name"] for c in e.get("columns", [])]
        extra = f" cols={cols[:4]}"
    elif t in ("heading", "subheading", "paragraph"):
        extra = f' text="{e.get("text","")[:60]}"'
    print(f"  {t}{extra}")

print("\n=== metadata (KV pairs) ===")
print(json.dumps(d["metadata"], indent=2))

print("\n=== content/sections summary (first 20) ===")
for s in d["content"][:20]:
    title = s.get("title", "")[:60]
    print(f"  Page {s.get('page')}: {s['type']} - {title}")

print(f"\n=== Total tables: {len(d['tables'])} ===")
for i, t in enumerate(d["tables"][:3]):
    print(f"  Table {i+1} (page {t['page']}): headers={t['headers'][:4]}, rows={len(t['rows'])}")

print(f"\n=== Total images: {len(d['images_and_graphs'])} ===")

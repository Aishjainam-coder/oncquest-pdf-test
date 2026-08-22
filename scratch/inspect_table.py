import json
from pathlib import Path

json_path = Path("extracted_jsons/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.json")
with open(json_path, "r", encoding="utf-8") as f:
    d = json.load(f)

for p_idx, p in enumerate(d.get("document", {}).get("pages", [])):
    for e_idx, e in enumerate(p.get("elements", [])):
        if e.get("type") == "table":
            cols = e.get("columns", [])
            has_w = any("width" in col if isinstance(col, dict) else False for col in cols)
            print(f"Page {p_idx} element {e_idx} table has {len(cols)} columns. Has width key: {has_w}")
            for col in cols:
                print("  ", col)

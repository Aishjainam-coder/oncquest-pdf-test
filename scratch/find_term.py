import json
from pathlib import Path

json_path = Path("extracted_jsons/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.json")
with open(json_path, "r", encoding="utf-8") as f:
    d = json.load(f)

def search(val, path=""):
    if isinstance(val, str):
        if "lunresertib" in val.lower():
            print(f"Found: {path} -> {val}")
    elif isinstance(val, dict):
        for k, v in val.items():
            search(v, f"{path}.{k}" if path else k)
    elif isinstance(val, list):
        for i, item in enumerate(val):
            search(item, f"{path}[{i}]")

search(d)

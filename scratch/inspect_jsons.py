import json
from pathlib import Path

json_dir = Path("extracted_jsons")
for p in json_dir.glob("*.json"):
    print("="*80)
    print(f"JSON File: {p.name}")
    print("="*80)
    try:
        with open(p, "r", encoding="utf-8") as f:
            data = json.load(f)
        
        # Look for metadata or keys related to test name
        print("Metadata:", json.dumps(data.get("metadata", {}), indent=2))
        
        # Search for occurrences of test names in the json keys/values
        print("Key-value pairs:")
        for k, v in data.get("extracted_key_value_pairs", {}).items():
            if "test" in k.lower() or "panel" in k.lower() or "name" in k.lower():
                print(f"  {k}: {v}")
                
        # Also check page 1 sections or headers
        print("Sections:")
        for sec in data.get("sections", []):
            if sec.get("page") == 1:
                title = sec.get("title", "")
                if title:
                    print(f"  Section Title: {title}")
    except Exception as e:
        print(f"Error reading {p.name}: {e}")
    print("\n")

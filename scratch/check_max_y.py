import json
import glob

max_y = 0
max_block = None
max_file = ""

for f in glob.glob("extracted_jsons/*.json"):
    with open(f, "r", encoding="utf-8") as file:
        d = json.load(file)
        for p in d.get("pages", []):
            for b in p.get("text_blocks", []):
                y1 = b["bbox"][3]
                # We want to see content below 650 but above 720 (where footer region starts)
                if y1 > max_y and y1 < 700:
                    max_y = y1
                    max_block = b
                    max_file = f

print("Max Y coordinate of body text blocks:")
print(f"Max Y: {max_y}")
print(f"Block: {max_block}")
print(f"File: {max_file}")

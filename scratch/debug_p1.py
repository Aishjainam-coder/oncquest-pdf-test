import json
with open("extracted_jsons/TestReport_SAROJ DEVI (OQG2604250052)_2600133939_1eb3f888-f97e-4181-b84d-d9536de7af26.json", encoding="utf-8") as f:
    data = json.load(f)
p1 = data["document"]["pages"][0]
for el in p1["elements"]:
    bbox = el.get("bbox")
    if bbox:
        y0 = bbox.get("y0", bbox.get("y", 0))
        if 320 <= y0 <= 420:
            print("Type:", el.get("type"), "y0:", round(y0, 1), "text:", repr(el.get("text", "") or el.get("title", "")))

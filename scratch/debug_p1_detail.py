import json
with open("extracted_jsons/TestReport_SAROJ DEVI (OQG2604250052)_2600133939_1eb3f888-f97e-4181-b84d-d9536de7af26.json", encoding="utf-8") as f:
    data = json.load(f)
p1 = data["document"]["pages"][0]
for el in p1["elements"]:
    if el.get("type") == "table" and "Gene & Transcript" in str(el.get("columns")):
        print("COLUMNS:", el.get("columns"))
        print("ROWS:", el.get("rows"))
        print("CELL_STYLES:", el.get("cell_styles"))

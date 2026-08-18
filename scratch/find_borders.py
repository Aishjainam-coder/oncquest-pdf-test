with open("converter.py", "r", encoding="utf-8") as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "border" in line.lower() or "table" in line.lower():
        if "css" in line.lower() or "style" in line.lower() or "border" in line.lower() or "def " in line.lower():
            print(f"Line {i+1}: {line.strip()[:100]}")

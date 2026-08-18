import re
from pathlib import Path

extractor_path = Path(r"c:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main\extractor.py")
with open(extractor_path, "r", encoding="utf-8") as f:
    content = f.read()

print("Functions matching 'detect_':")
for m in re.finditer(r"def\s+(\w+)\(", content):
    func_name = m.group(1)
    if "detect" in func_name or "bound" in func_name or "header" in func_name:
        # Find line number
        start = m.start()
        line_num = content[:start].count("\n") + 1
        print(f"  Line {line_num}: def {func_name}")

import re
from pathlib import Path

converter_path = Path(r"c:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main\converter.py")
with open(converter_path, "r", encoding="utf-8") as f:
    content = f.read()

# Let's locate the definition of convert_pdf_full_pipeline
match = re.search(r"def convert_pdf_full_pipeline\(.*?\):", content)
if match:
    start_pos = match.start()
    # Let's find the end of the function (until next 'def ' at the same indentation level)
    # The function is defined with indentation. Let's find where the indentation resets or next def starts
    lines = content[start_pos:].split("\n")
    func_lines = []
    for line in lines:
        if len(func_lines) > 0 and (line.startswith("def ") or line.startswith("class ")):
            break
        func_lines.append(line)
    
    print("\n".join(func_lines[:150]))  # Print first 150 lines of the function
else:
    print("Function convert_pdf_full_pipeline not found!")

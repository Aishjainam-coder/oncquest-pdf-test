"""
convert.py - Universal Dynamic CLI Runner
=========================================
Processes any input PDF document (lab report, invoice, certificate, bill, etc.),
extracts full structured JSON data (key-values, tables, content boxes, images, graphs),
and renders it into a dynamic HTML template with customizable design themes.
"""
import sys
import shutil
import json
from pathlib import Path
from converter import (
    process_pdf, 
    render_html_to_pdf_and_preview, 
    convert_json_to_docx, 
    convert_html_to_docx, 
    render_json_file_to_html,
    convert_pdf_full_pipeline
)
from extractor import extract_report_data

output_dir = Path("output")
output_dir.mkdir(exist_ok=True)
json_dir = Path("extracted_jsons")
json_dir.mkdir(exist_ok=True)

# Load theme settings from theme.json if present
theme_config = {}
theme_json_path = Path("theme.json")
if theme_json_path.exists():
    try:
        with open(theme_json_path, "r", encoding="utf-8") as f_theme:
            theme_config = json.load(f_theme)
    except Exception as e:
        print(f"[*] Note: Could not load theme.json: {e}")

# Determine input target file(s)
if len(sys.argv) > 1:
    target_path = Path(sys.argv[1])
    if target_path.exists():
        if target_path.suffix.lower() in [".html", ".htm"]:
            print(f"[*] Converting HTML file to Word (.docx): {target_path.name}")
            out_docx = output_dir / f"{target_path.stem}.docx"
            convert_html_to_docx(target_path, output_path=out_docx, theme_config=theme_config)
            print(f"[+] Successfully converted HTML -> Word (.docx): {out_docx}")
            sys.exit(0)
        elif target_path.suffix.lower() == ".json":
            print(f"[*] Rendering JSON file to HTML and Word (.docx): {target_path.name}")
            out_html = output_dir / f"{target_path.stem}.html"
            out_docx = output_dir / f"{target_path.stem}.docx"
            render_json_file_to_html(target_path, output_path=out_html, theme_config=theme_config)
            print(f"[+] Successfully rendered JSON -> HTML template: {out_html}")
            convert_html_to_docx(out_html, output_path=out_docx, theme_config=theme_config)
            print(f"[+] Successfully converted HTML -> Word (.docx) (exact HTML fidelity): {out_docx}")
            sys.exit(0)
        elif target_path.suffix.lower() == ".pdf":
            pdf_files = [target_path]
        else:
            print(f"[*] Error: Unsupported file format '{target_path.suffix}'. Expected .pdf, .json or .html")
            sys.exit(1)
    else:
        print(f"[*] Error: Specified file '{sys.argv[1]}' not found.")
        sys.exit(1)
else:
    # Find all PDF files in current directory
    pdf_files = [p for p in Path(".").glob("*.pdf") if p.is_file()]

if not pdf_files:
    print("[*] Warning: No PDF or HTML files found to process. Usage: python convert.py <path_to_pdf_or_html>")
    sys.exit(0)

print(f"[+] Found {len(pdf_files)} PDF file(s) to process dynamically: {[p.name for p in pdf_files]}\n")

SAVE_OUTPUT = True

for pdf_path in pdf_files:
    convert_pdf_full_pipeline(pdf_path, output_dir=output_dir, theme_config=theme_config)

print("\n[+] All 4-step dynamic conversions completed successfully!")


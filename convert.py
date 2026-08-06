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
from converter import process_pdf, render_html_to_pdf_and_preview, convert_json_to_docx, convert_html_to_docx, render_json_file_to_html
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
            theme_file_data = json.load(f_theme)
            colors = theme_file_data.get("colors", {})
            fonts = theme_file_data.get("fonts", {})
            theme_config = {
                "primary_color": colors.get("primary", "#1f497d"),
                "secondary_color": colors.get("secondary", "#008080"),
                "font_family": fonts.get("families", {}).get("primary", "Cambria, serif")
            }
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
            print(f"[*] Rendering JSON file to HTML: {target_path.name}")
            out_html = output_dir / f"{target_path.stem}.html"
            render_json_file_to_html(target_path, output_path=out_html, theme_config=theme_config)
            print(f"[+] Successfully rendered JSON -> HTML template: {out_html}")
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
    pdf_stem = pdf_path.stem
    print(f"[*] Processing input PDF dynamically: {pdf_path.name}")
    
    # 1. Universal Extraction
    extracted_data = extract_report_data(str(pdf_path))
    if extracted_data:
        print(f"   [+] Data Extracted Successfully for: {pdf_path.name}")
        print(f"        - Key-Value Pairs: {len(extracted_data.get('all_key_value_pairs', {}))}")
        print(f"        - Tables: {len(extracted_data.get('all_tables', []))}")
        print(f"        - Content Boxes: {len(extracted_data.get('all_boxes_and_sections', []))}")
        print(f"        - Images & Graphs: {len(extracted_data.get('all_images_and_graphs', []))}")

    # 2. Render HTML
    out_html = output_dir / f"{pdf_stem}_target.html"
    full_html = process_pdf(str(pdf_path), out_html, is_target=False, use_template=False, theme_config=theme_config, save_output=SAVE_OUTPUT)
    print(f"   [+] Rendered HTML saved: {out_html}")
    
    # 3. Convert JSON -> Word (.docx)
    out_docx = output_dir / f"{pdf_stem}_report.docx"
    if extracted_data:
        convert_json_to_docx(extracted_data, output_path=out_docx, theme_config=theme_config)
    else:
        convert_html_to_docx(full_html, output_path=out_docx, theme_config=theme_config)
    print(f"   [+] Converted JSON to Word (.docx) with clean alignment: {out_docx}")

print("\n[+] Dynamic conversions completed successfully!")


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
from converter import process_pdf, render_html_to_pdf_and_preview, convert_json_to_docx
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

# 1. Determine which PDF files to process
pdf_files = []
if len(sys.argv) > 1:
    target_path = Path(sys.argv[1])
    if target_path.exists() and target_path.suffix.lower() == ".pdf":
        pdf_files.append(target_path)
    else:
        print(f"[*] Error: Specified file '{sys.argv[1]}' not found or is not a PDF.")
        sys.exit(1)
else:
    # Find all PDF files in current directory (excluding output/ folder)
    pdf_files = [p for p in Path(".").glob("*.pdf") if p.is_file()]

if not pdf_files:
    print("[*] Warning: No PDF files found to process. Usage: python convert.py <path_to_any_pdf>")
    sys.exit(0)

print(f"[+] Found {len(pdf_files)} PDF file(s) to process dynamically: {[p.name for p in pdf_files]}\n")

# Flag to control saving files to the output/ folder (disabled as requested)
SAVE_OUTPUT = False

for pdf_path in pdf_files:
    pdf_stem = pdf_path.stem
    print(f"[*] Processing input PDF dynamically: {pdf_path.name}")
    
    # 2. Extract Structured JSON Data (in-memory)
    extracted_data = extract_report_data(str(pdf_path))
    if extracted_data:
        print(f"   [+] Data Extracted Successfully for: {pdf_path.name}")
        print(f"        - Key-Value Pairs: {len(extracted_data.get('all_key_value_pairs', {}))}")
        print(f"        - Tables: {len(extracted_data.get('all_tables', []))}")
        print(f"        - Content Boxes: {len(extracted_data.get('all_boxes_and_sections', []))}")
        print(f"        - Images & Graphs: {len(extracted_data.get('all_images_and_graphs', []))}")

    # 3. Process HTML (in-memory, output file saving disabled)
    out_html = output_dir / f"{pdf_stem}_target.html" if SAVE_OUTPUT else None
    full_html = process_pdf(str(pdf_path), out_html, is_target=False, use_template=False, theme_config=theme_config, save_output=SAVE_OUTPUT)
    
    if SAVE_OUTPUT:
        out_pdf = output_dir / f"{pdf_stem}_output.pdf"
        out_img = output_dir / f"{pdf_stem}_preview.png"
        render_html_to_pdf_and_preview(out_html, out_pdf, out_img)
        print(f"   [->] Saved to {output_dir}/ folder")
    else:
        print("   [+] Processing completed in-memory (saving to output folder is disabled).")

print("\n[+] Dynamic conversions completed in-memory!")

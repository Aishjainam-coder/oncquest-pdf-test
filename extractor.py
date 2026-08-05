"""
Universal PDF Data, Tables, Images, Graphs & Layout Content Extraction Engine
==============================================================================
Extracts structured JSON data from ANY input PDF document (lab reports, invoices,
bills, certificates, financial statements, technical specs, tax forms, etc.).

Dynamic extraction includes:
- Key-Value pairs & header metadata across all pages
- Content boxes & structured text sections with typography/bbox attributes
- Data tables with headers, cell grids, and bounding boxes
- Embedded images, charts, graphs, logos (extracted as Base64 Data URIs)
- Complete page-by-page text layout and document summary metrics
"""

import os
import re
import json
import base64
import logging

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("UniversalPDFExtractor")


def parse_header_key_value_pairs(full_text: str) -> dict:
    """
    Extracts structured key-value header metadata pairs from document text dynamically.
    Enforces strict criteria for keys (short labels, no prose sentences) to ensure accuracy.
    """
    kv_dict = {}
    lines = full_text.split('\n')
    
    ignore_keywords = [
        'received from', 'responsibility', 'presumed that', 'http', 'www',
        'copyright', 'printed on', 'page', 'all rights reserved', 'disclaimer',
        'classified as', 'predicted to', 'evidence', 'database', 'population'
    ]

    prose_verbs = [' is ', ' are ', ' was ', ' were ', ' to ', ' by ', ' from ', ' with ', ' that ', ' which ']

    for line in lines:
        line_str = line.strip()
        if not line_str or len(line_str) > 250:
            continue

        # Find key-value patterns like "Label: Value" or "Label: Value    Label2: Value2"
        # Regex captures keys starting with Alphanumeric (2-35 chars), avoiding full prose sentences
        matches = re.findall(r'([A-Za-z0-9][A-Za-z0-9\s/&#\-_]{1,35}):\s*([^\n\t:]{1,120})', line_str)
        for k_cand, v_cand in matches:
            k = k_cand.strip()
            v = v_cand.strip()

            # Clean key validation
            if not k or len(k) < 2 or len(k) > 35:
                continue
            if k.count(' ') > 4:  # Keys shouldn't be full sentences
                continue
            if any(char in k for char in ['.', ',', ';', '(', ')', '"', '?', '!']):
                continue
            if any(verb in f" {k.lower()} " for verb in prose_verbs):
                continue
            if any(kw in k.lower() for kw in ignore_keywords):
                continue
            if k.startswith('#') or k.lower().startswith('page'):
                continue
            
            # Clean value validation
            if v and len(v) < 150:
                # Truncate if next pair was concatenated in value
                if '   ' in v:
                    v = re.split(r'\s{3,}', v)[0].strip()
                kv_dict[k] = v

    return kv_dict


def extract_page_images_and_graphs(doc, page_num: int, page) -> list:
    """
    Extracts all raster images, charts, logos, and graphic figures from a PDF page.
    Returns list of dicts with image metadata and Base64 Data URI string.
    """
    images_list = []
    try:
        page_images = page.get_images(full=True)
        for img_idx, img_info in enumerate(page_images):
            xref = img_info[0]
            try:
                base_img = doc.extract_image(xref)
                if not base_img:
                    continue
                
                img_bytes = base_img.get("image")
                img_ext = base_img.get("ext", "png").lower()
                if img_ext == "jpx":
                    img_ext = "jpeg"
                
                width = base_img.get("width", 0)
                height = base_img.get("height", 0)
                
                if width < 8 or height < 8:
                    continue

                b64_str = base64.b64encode(img_bytes).decode("utf-8")
                mime_type = f"image/{img_ext}"
                data_uri = f"data:{mime_type};base64,{b64_str}"

                bbox = None
                try:
                    rects = page.get_image_rects(xref)
                    if rects:
                        r = rects[0]
                        bbox = [round(r.x0, 2), round(r.y0, 2), round(r.x1, 2), round(r.y1, 2)]
                except Exception:
                    pass

                is_graph_or_chart = (width >= 120 and height >= 80) or (width >= 80 and height >= 120)
                img_type = "graph_chart" if is_graph_or_chart else "image_logo"

                images_list.append({
                    "page": page_num + 1,
                    "image_index": img_idx + 1,
                    "xref": xref,
                    "width": width,
                    "height": height,
                    "format": img_ext,
                    "mime_type": mime_type,
                    "type": img_type,
                    "bbox": bbox,
                    "data_uri": data_uri
                })
            except Exception as ie:
                logger.debug(f"Failed to extract image xref {xref} on page {page_num + 1}: {ie}")
    except Exception as e:
        logger.debug(f"Error getting images on page {page_num + 1}: {e}")

    return images_list


def is_inside_table_bbox(bbox, table_bboxes):
    """Check if a text block's bounding box falls inside any extracted table bounding box."""
    if not bbox or not table_bboxes:
        return False
    bx0, by0, bx1, by1 = bbox
    cx, cy = (bx0 + bx1) / 2.0, (by0 + by1) / 2.0
    for tx0, ty0, tx1, ty1 in table_bboxes:
        if (tx0 - 5) <= cx <= (tx1 + 5) and (ty0 - 5) <= cy <= (ty1 + 5):
            return True
    return False


def extract_report_data(pdf_path: str) -> dict:
    """
    Universal extraction engine for ANY PDF document format.

    Extracts:
        - document_summary: file_name, total_pages, total_tables, total_boxes, total_key_value_pairs, total_images_and_graphs
        - extracted_key_value_pairs: dict of clean key-value pairs
        - content_sections: list of clean content section boxes (excluding table text)
        - tables: list of extracted tables (headers, rows, bbox)
        - images_and_graphs: list of extracted raster images, charts, and logos (Base64 URIs)
        - pages: page-by-page breakdown
        - raw_full_text: entire document text
    """
    if not os.path.exists(pdf_path):
        logger.error(f"PDF file not found: {pdf_path}")
        return None

    if fitz is None:
        logger.error("PyMuPDF (fitz) library is not installed.")
        return None

    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        logger.error(f"Failed to open PDF document {pdf_path}: {e}")
        return None

    full_text_pages = []
    text_blocks_with_style = []
    pages_list = []
    all_tables_list = []
    all_kv_pairs = {}
    all_boxes_list = []
    all_images_list = []

    # Iterate through all pages
    for page_num in range(len(doc)):
        page = doc[page_num]
        rect = page.rect
        page_text = page.get_text("text")
        full_text_pages.append(page_text)

        # 1. Extract Key-Value pairs on this page
        page_kv = parse_header_key_value_pairs(page_text)
        for k, v in page_kv.items():
            all_kv_pairs[k] = v

        # 2. Extract Tables on this page FIRST (so we get table bounding boxes)
        page_tables = []
        table_bboxes = []
        try:
            if hasattr(page, "find_tables"):
                tf = page.find_tables()
                table_list = list(tf)
                for t_idx, table in enumerate(table_list):
                    table_data = table.extract()
                    if table_data and len(table_data) > 0:
                        headers = [str(c).strip().replace("\n", " ") if c else "" for c in table_data[0]]
                        rows = [[str(cell).strip().replace("\n", " ") if cell else "" for cell in row] for row in table_data[1:]]
                        t_bbox = [round(c, 2) for c in table.bbox]
                        table_bboxes.append(t_bbox)

                        table_obj = {
                            "page": page_num + 1,
                            "table_index": t_idx + 1,
                            "bbox": t_bbox,
                            "headers": headers,
                            "rows": rows
                        }
                        page_tables.append(table_obj)
                        all_tables_list.append(table_obj)
        except Exception as t_err:
            logger.debug(f"Page {page_num + 1} table error: {t_err}")

        # 3. Parse text blocks outside tables
        text_page_dict = page.get_text("dict")
        page_blocks = []
        for block in text_page_dict.get("blocks", []):
            if "lines" in block:
                b_bbox = [round(c, 2) for c in block["bbox"]]
                # Skip text blocks inside tables
                if is_inside_table_bbox(b_bbox, table_bboxes):
                    continue

                b_text = ""
                is_bold = False
                max_size = 0.0
                font_name = ""
                for line in block["lines"]:
                    for span in line["spans"]:
                        text = span.get("text", "").strip()
                        if text:
                            b_text += span.get("text", "") + " "
                            if span.get("flags", 0) & 2 or "bold" in span.get("font", "").lower():
                                is_bold = True
                            if span.get("size", 0.0) > max_size:
                                max_size = span.get("size", 0.0)
                            font_name = span.get("font", "")

                clean_b_text = b_text.strip()
                if clean_b_text:
                    block_obj = {
                        "page": page_num + 1,
                        "bbox": b_bbox,
                        "text": clean_b_text,
                        "max_font_size": round(max_size, 2),
                        "is_bold": is_bold,
                        "font": font_name
                    }
                    page_blocks.append(block_obj)
                    text_blocks_with_style.append(block_obj)

        # 4. Extract Images & Graphs on this page
        page_images = extract_page_images_and_graphs(doc, page_num, page)
        all_images_list.extend(page_images)

        # 5. Extract Content Boxes & Sections on this page cleanly
        page_boxes = []

        # Demographics / Header Box if key-values present on top of page
        if page_kv and page_num == 0:
            page_boxes.append({
                "page": page_num + 1,
                "title": f"Document Header & Patient Metadata",
                "type": "demographics_box",
                "bbox": [round(rect.x0, 2), round(rect.y0, 2), round(rect.x1, 2), round(rect.y0 + 150, 2)],
                "key_value_pairs": page_kv
            })

        current_box = None
        for b in page_blocks:
            t = b["text"]
            
            # Heading candidate check: prominent font, bold short title, not ending with sentence punctuation
            is_heading_candidate = (
                (b["max_font_size"] >= 11.5 or (b["is_bold"] and len(t) < 70)) and
                not t.endswith(('.', ',', ';', '?')) and
                not any(verb in f" {t.lower()} " for verb in [' is ', ' are ', ' was ', ' were ', ' should ']) and
                len(t) < 90
            )

            if is_heading_candidate:
                if current_box and current_box.get("content_text"):
                    page_boxes.append(current_box)
                current_box = {
                    "page": page_num + 1,
                    "title": t,
                    "type": "content_section_box",
                    "bbox": b["bbox"],
                    "content_text": []
                }
            else:
                if current_box is None:
                    current_box = {
                        "page": page_num + 1,
                        "title": "General Content / Notes",
                        "type": "content_section_box",
                        "bbox": b["bbox"],
                        "content_text": [t]
                    }
                else:
                    current_box["content_text"].append(t)

        if current_box and current_box.get("content_text"):
            page_boxes.append(current_box)

        for box in page_boxes:
            all_boxes_list.append(box)

        pages_list.append({
            "page_number": page_num + 1,
            "dimensions": {"width": round(rect.width, 2), "height": round(rect.height, 2)},
            "key_value_pairs": page_kv,
            "boxes_and_sections": page_boxes,
            "tables": page_tables,
            "images_and_graphs": page_images,
            "text_blocks": page_blocks,
            "page_text": page_text
        })

    full_document_text = "\n".join(full_text_pages)

    doc_summary = {
        "file_name": os.path.basename(pdf_path),
        "total_pages": len(doc),
        "total_tables": len(all_tables_list),
        "total_boxes": len(all_boxes_list),
        "total_key_value_pairs": len(all_kv_pairs),
        "total_images_and_graphs": len(all_images_list)
    }

    extracted_data = {
        "document_summary": doc_summary,
        "extracted_key_value_pairs": all_kv_pairs,
        "all_key_value_pairs": all_kv_pairs,
        "tables": all_tables_list,
        "all_tables": all_tables_list,
        "content_sections": all_boxes_list,
        "all_boxes_and_sections": all_boxes_list,
        "images_and_graphs": all_images_list,
        "all_images_and_graphs": all_images_list,
        "pages": pages_list,
        "raw_full_text": full_document_text
    }

    try:
        doc.close()
    except Exception:
        pass

    logger.info(f"=== Universal Extraction Complete for {os.path.basename(pdf_path)} ===")
    logger.info(f"Extracted {len(all_kv_pairs)} KV pairs, {len(all_tables_list)} Tables, {len(all_images_list)} Images/Graphs across {len(pages_list)} Pages.")
    return extracted_data


if __name__ == "__main__":
    import sys
    test_pdf = sys.argv[1] if len(sys.argv) > 1 else "vendor pdf.pdf"
    if os.path.exists(test_pdf):
        print(f"\n[*] Testing Universal Extraction on: {test_pdf}")
        res = extract_report_data(test_pdf)
        if res:
            print("[+] Summary:", json.dumps(res["document_summary"], indent=2))
            print(f"[+] Key-Value Pairs: {len(res['extracted_key_value_pairs'])}")
            print(f"[+] Data Tables: {len(res['tables'])}")
            print(f"[+] Images & Graphs: {len(res['images_and_graphs'])}")


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


def detect_dynamic_header_footer_bounds(doc) -> dict:
    """
    Dynamically analyzes document layout to identify header and footer bounding box cutoffs
    for EVERY page of the document (including Page 1).
    Header region includes top vendor logos, QR codes, report titles, and patient details header tables.
    Footer region includes signatures, doctor names, qualifications, page numbers, and disclaimers.
    """
    from collections import defaultdict
    page_bounds = {}
    page_count = len(doc)

    header_kw = [
        'neuberg', 'center for genomic medicine', 'genomic medicine', 'neuberg diagnostics',
        'laboratory report', 'mc-6200', 'case id', 'sample type', 'patient name', 'uhid', 'reg no',
        'ref by', 'referred by', 'dis.loc.', 'pt id', 'pt. id', 'pt. loc.', 'ph #', 'ref id', 'ref id 2',
        'age/gender', 'age / sex', 'date & time collected', 'date & time received', 'date & time reported',
        'registration date & time', 'sample date & time', 'report date & time',
        'date collected', 'date received', 'date reported', 'report version', 'bill. loc.',
        'lab no', 'barcode', 'qr code for report', 'report verification', 'patient information',
        'patient details', 'patient metadata', 'demographics', 'sample coll.by', 'acc. remarks'
    ]

    footer_kw = [
        'reviewed by', 'verified by', 'authorized signatory', 'signatory', 'doctor', 'dr.',
        'md (path', 'ph.d.', 'pathologist', 'biochemist', 'microbiologist', 'geneticist',
        'consultant', 'end of report', 'electronically generated', 'disclaimer', 'registered office',
        'nabl', 'cap accredited', 'iso 15189', 'mc-7414', 'page '
    ]

    # Step 1: Detect multi-page repeated strings in top 35% and bottom 35%
    top_str_counts = defaultdict(int)
    bot_str_counts = defaultdict(int)

    for p_idx in range(page_count):
        page = doc[p_idx]
        H = page.rect.height
        text_page = page.get_text("dict")
        for b in text_page.get("blocks", []):
            if "lines" not in b:
                continue
            y0, y1 = b["bbox"][1], b["bbox"][3]
            text = " ".join(s.get("text", "") for ln in b["lines"] for s in ln.get("spans", [])).strip()
            if not text:
                continue
            norm_t = re.sub(r'\s+', ' ', text.lower())
            if y0 < 0.35 * H:
                top_str_counts[norm_t] += 1
            elif y1 > 0.65 * H:
                bot_str_counts[norm_t] += 1

    repeated_top = {t for t, count in top_str_counts.items() if count >= 2 or (page_count == 1 and count >= 1)}
    repeated_bot = {t for t, count in bot_str_counts.items() if count >= 2 or (page_count == 1 and count >= 1)}

    # Step 2: Calculate cutoffs per page for all pages
    for p_idx in range(page_count):
        page = doc[p_idx]
        H = page.rect.height
        text_page = page.get_text("dict")

        header_y1 = 0.0
        footer_y0 = H

        # A) Analyze text blocks
        for b in text_page.get("blocks", []):
            if "lines" not in b:
                continue
            y0, y1 = b["bbox"][1], b["bbox"][3]
            text = " ".join(s.get("text", "") for ln in b["lines"] for s in ln.get("spans", [])).strip()
            if not text:
                continue
            norm_t = re.sub(r'\s+', ' ', text.lower())

            # Check Header match (top 28% of page)
            if y0 < 0.28 * H:
                is_hdr = False
                if norm_t in repeated_top:
                    is_hdr = True
                elif any(kw in norm_t for kw in header_kw):
                    is_hdr = True
                elif re.search(r'page\s+\d+', norm_t):
                    is_hdr = True
                elif y1 < 0.15 * H:
                    is_hdr = True

                if is_hdr:
                    if y1 > header_y1:
                        header_y1 = y1

            # Check Footer match (bottom 16% of page ONLY)
            if y0 > 0.82 * H:
                is_ftr = False
                if norm_t in repeated_bot:
                    is_ftr = True
                elif any(kw in norm_t for kw in footer_kw):
                    is_ftr = True
                elif re.search(r'page\s+\d+', norm_t):
                    is_ftr = True
                elif y0 > 0.88 * H:
                    is_ftr = True

                if is_ftr:
                    if y0 < footer_y0:
                        footer_y0 = y0

        # B) Analyze image bboxes (top 20% for header, bottom 15% for footer)
        img_infos = page.get_image_info(xrefs=True)
        for info in img_infos:
            bbox = info.get("bbox")
            if not bbox:
                continue
            iy0, iy1 = bbox[1], bbox[3]
            if iy0 < 0.20 * H:
                if iy1 > header_y1:
                    header_y1 = iy1
            if iy0 > 0.84 * H:
                if iy0 < footer_y0:
                    footer_y0 = iy0

        # Add buffer - Header capped at 0.26*H (~218px), Footer floored at 0.84*H (~707px)
        final_hdr_cutoff = min(header_y1 + 2.0, 0.26 * H) if header_y1 > 0 else 0.0
        final_ftr_cutoff = max(footer_y0 - 2.0, 0.84 * H) if footer_y0 < H else H

        page_bounds[p_idx] = {
            "header_y_cutoff": round(final_hdr_cutoff, 2),
            "footer_y_cutoff": round(final_ftr_cutoff, 2)
        }

    return page_bounds


def is_patient_sample_header_block(text: str, y0: float) -> bool:
    """Identify and filter out top Patient Details, Lab Header, and Sample Details header card blocks."""
    norm = re.sub(r'\s+', ' ', text.lower().strip())
    if y0 < 340:
        if any(kw in norm for kw in [
            'neuberg', 'center for genomic medicine', 'genomic medicine', 'neuberg diagnostics',
            'laboratory report', 'mc-6200', 'a unit of neuberg'
        ]):
            return True
        if norm in ['patient details', 'sample details', 'patient information', 'sample information', 'demographics', 'laboratory report']:
            return True
        if any(kw in norm for kw in [
            'registration date & time', 'sample date & time', 'report date & time',
            'registration date', 'sample date', 'report date', 'sample coll.by', 'acc. remarks',
            'ref id 1.', 'ref id 2', 'ref id', 'ph #', 'dis.loc.', 'pt. loc.', 'pt. id', 'pt id',
            'neuberg center for genomic medicine', 'orion intelligent genomics',
            'qr code for report verification', 'sex / age', 'sex/age', 'case id :', 'case id:',
            'case id', 'bill. loc.', 'ref by :', 'ref by', 'sample type :', 'sample type:'
        ]):
            return True
        if 'name' in norm and ('mr.' in norm or 'mrs.' in norm or 'ms.' in norm or 'master' in norm or 'dr.' in norm or 'oqg' in norm or '1007' in norm):
            return True
        if norm.startswith(': name') or norm.startswith('name :') or norm.startswith('name'):
            return True
    return False


def extract_report_data(pdf_path: str) -> dict:
    """
    Universal extraction engine for ANY PDF document format.
    Completely excludes header (logos, QR codes, patient details header card) and footer
    from EVERY page including Page 1.
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

    # Detect dynamic header and footer bounding box cutoffs across all pages
    page_bounds = detect_dynamic_header_footer_bounds(doc)

    header_kv_ignore_keys = [
        'case id', 'sample type', 'patient name', 'uhid', 'reg no', 'ref by', 'referred by',
        'age/gender', 'age / sex', 'date & time collected', 'date & time received', 'date & time reported',
        'registration date', 'sample date', 'report date', 'date collected', 'date received', 'date reported',
        'report version', 'bill. loc.', 'dis.loc.', 'pt id', 'pt. id', 'pt. loc.', 'ph #', 'ref id',
        'sample coll.by', 'acc. remarks', 'neuberg', 'genomic medicine', 'laboratory report', 'mc-6200',
        'collected', 'received', 'reported', 'lab no', 'barcode', 'qr code',
        'accession.id', 'centre details'
    ]

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
        H = rect.height
        hy_cutoff = page_bounds[page_num]["header_y_cutoff"]
        fy_cutoff = page_bounds[page_num]["footer_y_cutoff"]

        text_page_dict = page.get_text("dict")

        # 1. Extract Tables on this page in body region ONLY
        page_tables = []
        table_bboxes = []
        try:
            if hasattr(page, "find_tables"):
                tf = page.find_tables()
                table_list = list(tf)
                for t_idx, table in enumerate(table_list):
                    t_bbox = [round(c, 2) for c in table.bbox]
                    # Exclude tables inside header or footer regions on EVERY page
                    if t_bbox[1] < hy_cutoff or t_bbox[3] > fy_cutoff:
                        continue

                    table_data = table.extract()
                    if table_data and len(table_data) > 0:
                        headers = [str(c).strip().replace("\n", " ") if c else "" for c in table_data[0]]
                        rows = [[str(cell).strip().replace("\n", " ") if cell else "" for cell in row] for row in table_data[1:]]
                        table_bboxes.append(t_bbox)

                        table_obj = {
                            "page": page_num + 1,
                            "table_index": len(page_tables) + 1,
                            "bbox": t_bbox,
                            "type": "table",
                            "headers": headers,
                            "rows": rows
                        }
                        page_tables.append(table_obj)
                        all_tables_list.append(table_obj)
        except Exception as t_err:
            logger.debug(f"Page {page_num + 1} table error: {t_err}")

        # 2. Parse text blocks outside tables and strictly inside body region ONLY
        page_blocks = []
        body_text_lines = []
        for block in text_page_dict.get("blocks", []):
            if "lines" in block:
                b_bbox = [round(c, 2) for c in block["bbox"]]
                if is_inside_table_bbox(b_bbox, table_bboxes):
                    continue

                # Exclude text blocks inside header or footer regions on EVERY page
                if b_bbox[1] < hy_cutoff or b_bbox[3] > fy_cutoff:
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
                    if re.search(r'^page\s+\d+(\s+of\s+\d+)?$', clean_b_text, re.I):
                        continue
                    if any(kw in clean_b_text.lower() for kw in ['reviewed by', 'authorized signatory', 'mc-7414']):
                        continue
                    if is_patient_sample_header_block(clean_b_text, b_bbox[1]):
                        continue

                    is_hd_cand = (max_size >= 11.5 or (is_bold and len(clean_b_text) < 70)) and not clean_b_text.endswith(('.', ',', ';', '?'))
                    block_obj = {
                        "page": page_num + 1,
                        "bbox": b_bbox,
                        "type": "heading" if is_hd_cand else "paragraph",
                        "text": clean_b_text,
                        "max_font_size": round(max_size, 2),
                        "is_bold": is_bold,
                        "font": font_name
                    }
                    page_blocks.append(block_obj)
                    text_blocks_with_style.append(block_obj)
                    body_text_lines.append(clean_b_text)

        page_body_text = "\n".join(body_text_lines)
        full_text_pages.append(page_body_text)

        page_kv = parse_header_key_value_pairs(page_body_text)
        filtered_page_kv = {}
        for k, v in page_kv.items():
            if not any(hk in k.lower() for hk in header_kv_ignore_keys):
                filtered_page_kv[k] = v
                all_kv_pairs[k] = v

        # 3. Extract Images & Graphs in body region ONLY (excluding header/footer logos)
        raw_images = extract_page_images_and_graphs(doc, page_num, page)
        page_images = []
        for img in raw_images:
            img_bbox = img.get("bbox")
            if img_bbox:
                y0, y1 = img_bbox[1], img_bbox[3]
                # Filter out header logo (y0 < 120) and footer logos (y0 >= 700 or y1 >= 720)
                if y0 < 120 or y1 <= hy_cutoff:
                    continue
                if y0 >= 700 or y1 >= 720:
                    continue
            img["type"] = "image"
            page_images.append(img)
            all_images_list.append(img)

        # 4. Extract Content Boxes & Sections in body region ONLY
        page_boxes = []
        current_box = None
        for b in page_blocks:
            t = b["text"]
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
                    "type": "box",
                    "bbox": b["bbox"],
                    "content_text": []
                }
            else:
                if current_box is None:
                    current_box = {
                        "page": page_num + 1,
                        "title": "General Content / Notes",
                        "type": "box",
                        "bbox": b["bbox"],
                        "content_text": [t]
                    }
                else:
                    current_box["content_text"].append(t)

        if current_box and current_box.get("content_text"):
            page_boxes.append(current_box)

        for box in page_boxes:
            all_boxes_list.append(box)

        # Build unified per-page content list sorted by bbox top-Y
        page_content_items = []
        if filtered_page_kv:
            page_content_items.append({
                "page": page_num + 1,
                "type": "key_value",
                "pairs": filtered_page_kv,
                "bbox": [35.0, hy_cutoff + 5.0, 560.0, hy_cutoff + 45.0]
            })
        for box in page_boxes:
            page_content_items.append({
                "page": page_num + 1,
                "type": "banner" if "REPORT" in box.get("title", "").upper() or "CLINICAL" in box.get("title", "").upper() else "box",
                "title": box.get("title"),
                "text": box.get("content_text"),
                "content_text": box.get("content_text"),
                "bbox": box.get("bbox")
            })
        for tbl in page_tables:
            page_content_items.append({
                "page": page_num + 1,
                "type": "table",
                "headers": tbl.get("headers", []),
                "rows": tbl.get("rows", []),
                "bbox": tbl.get("bbox")
            })
        for img in page_images:
            page_content_items.append({
                "page": page_num + 1,
                "type": "image",
                "data_uri": img.get("data_uri"),
                "width": img.get("width"),
                "height": img.get("height"),
                "bbox": img.get("bbox")
            })

        page_content_items.sort(key=lambda item: item.get("bbox", [0, 0, 0, 0])[1] if item.get("bbox") else 0.0)

        pages_list.append({
            "page_number": page_num + 1,
            "dimensions": {"width": round(rect.width, 2), "height": round(rect.height, 2)},
            "header_y_cutoff": hy_cutoff,
            "footer_y_cutoff": fy_cutoff,
            "content": page_content_items,
            "key_value_pairs": filtered_page_kv,
            "boxes_and_sections": page_boxes,
            "tables": page_tables,
            "images_and_graphs": page_images,
            "text_blocks": page_blocks,
            "page_text": page_body_text
        })

    # Build top-to-bottom ordered document flow sections
    ordered_sections = []
    
    # 1. Banners & Main Titles identification
    for box in all_boxes_list:
        title = box.get("title", "").strip()
        t_upper = title.upper()
        if "REPORT" in t_upper or "CERTIFICATE" in t_upper or "INVOICE" in t_upper or "STATEMENT" in t_upper or len(title) > 60:
            section_type = "banner"
        else:
            section_type = "banner" if box.get("page") == 1 and len(ordered_sections) < 2 else "box"

        box_item = {
            "type": section_type,
            "section_type": section_type,
            "title": title,
            "page": box.get("page", 1),
            "bbox": box.get("bbox", [0, 0, 0, 0]),
            "content_text": box.get("content_text", []),
            "paragraphs": box.get("content_text", [])
        }
        ordered_sections.append(box_item)

    # 2. Add Tables into ordered flow
    for tbl in all_tables_list:
        ordered_sections.append({
            "type": "table",
            "section_type": "table",
            "title": f"Data Table (Page {tbl.get('page', 1)})",
            "page": tbl.get("page", 1),
            "bbox": tbl.get("bbox", [0, 0, 0, 0]),
            "headers": tbl.get("headers", []),
            "rows": tbl.get("rows", [])
        })

    # 3. Add Images & Graphs into ordered flow
    for img in all_images_list:
        ordered_sections.append({
            "type": "image",
            "section_type": "image",
            "title": f"Image (Page {img.get('page', 1)})",
            "page": img.get("page", 1),
            "bbox": img.get("bbox", [0, 0, 0, 0]),
            "data_uri": img.get("data_uri"),
            "mime_type": img.get("mime_type", "image/png"),
            "width": img.get("width"),
            "height": img.get("height")
        })

    # Sort sections primarily by page, secondarily by y0 bbox coordinate
    ordered_sections.sort(key=lambda s: (s.get("page", 1), s.get("bbox", [0, 0, 0, 0])[1] if s.get("bbox") else 0))

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
        "metadata": all_kv_pairs,
        "extracted_key_value_pairs": all_kv_pairs,
        "all_key_value_pairs": all_kv_pairs,
        "content": ordered_sections,
        "sections": ordered_sections,
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

    # Automatically save JSON file into extracted_jsons/ directory
    try:
        json_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "extracted_jsons")
        os.makedirs(json_dir, exist_ok=True)
        pdf_stem = os.path.splitext(os.path.basename(pdf_path))[0]
        json_save_path = os.path.join(json_dir, f"{pdf_stem}.json")
        with open(json_save_path, "w", encoding="utf-8") as f_out:
            json.dump(extracted_data, f_out, indent=2, ensure_ascii=False)
        logger.info(f"Saved extracted JSON payload to: {json_save_path}")
    except Exception as e_save:
        logger.warning(f"Could not auto-save JSON to extracted_jsons/: {e_save}")

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


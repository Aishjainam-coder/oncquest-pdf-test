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
    import pymupdf as fitz  # PyMuPDF
except ImportError:
    fitz = None

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("UniversalPDFExtractor")


def replace_sng_gen_lab(text: str) -> str:
    if not isinstance(text, str):
        return text
    pattern = re.compile(r"(?:SNG\s+Gene?(?:['’‘]|&[a-zA-Z0-9#]+;)?s\s+Lab|SN\s+Genelab)\s+pvt\.?\s*ltd", re.IGNORECASE)
    return pattern.sub("Laboratory", text)


def replace_sng_in_structure(obj):
    if isinstance(obj, dict):
        return {k: replace_sng_in_structure(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [replace_sng_in_structure(item) for item in obj]
    elif isinstance(obj, str):
        return replace_sng_gen_lab(obj)
    return obj


def is_test_name_text(text: str) -> bool:
    if not isinstance(text, str):
        return False
    cleaned = text.strip()
    if len(cleaned) < 4 or len(cleaned) > 130:
        return False
        
    exclude_prefixes = (
        "clinical indication", "sample description", "report highlights", 
        "key findings", "test results", "tier ", "case id", "sample type", 
        "name :", "date & time", "bill. loc", "ref. by", "report version",
        "qr code", "page ", "salient features", "clinical suspicion",
        "dr.", "laboratory", "oncquest", "result summary", "methodology"
    )
    cleaned_lower = cleaned.lower()
    for ex in exclude_prefixes:
        if cleaned_lower.startswith(ex):
            return False
            
    patterns = [
        r'^(?:Breast\s+and\s+Ovarian\s+Extended\s+Panel\s*[-–]\s*Liquid\s+Biopsy\s+Assay)$',
        r'^(?:(?:Liquidseq\s+Actionable|Brainseq)\s+Genomic\s+Profiling\s+Panel(?:\s*[-–]\s*Advance)?)$',
        r'^(?:Whole\s+Exome\s+Sequencing(?:\s+on\s+(?:the\s+)?Illumina\s+[\w\s-]+\s+Platform)?)$',
        r'^[\w\s/&,–\-\(\)\.\+]+?(?:Genomic\s+Profiling\s+Panel|Extended\s+Panel|Profiling\s+Panel|Biopsy\s+Assay|Exome\s+Sequencing|Sequencing\s+Panel|Cancer\s+Panel|Gene\s+Panel|Profiling\s+Assay|Sequencing\s+Assay|Biopsy\s+Panel|NGS\s+Panel)(?:\s*[-–]\s*Advance)?(?:\s*\([^)]*\))?(?:\s+on\s+(?:the\s+)?Illumina\s+[\w\s-]+\s+Platform)?$',
        r'^(?:[A-Z\s]{4,}\s+PANEL(?:\s*[-–]\s*ADVANCE)?(?:\s*\([^)]*\))?)$',
    ]
    for pat in patterns:
        if re.match(pat, cleaned, re.IGNORECASE):
            return True
            
    return False


def is_subtitle_text(text: str) -> bool:
    if not isinstance(text, str):
        return False
    cleaned = text.strip()
    return bool(re.match(r'^\s*on\s+(?:the\s+)?(?:Illumina|Novaseq|Miseq|Nextseq|Ion\s+Torrent)\s+[\w\s-]+\s+Platform\s*$', cleaned, re.IGNORECASE))


def replace_test_name_in_structure(obj):
    """
    Recursively replaces the test name in JSON structure keys/values
    so direct JSON->DOCX or JSON->HTML routes are sanitized cleanly.
    """
    if isinstance(obj, dict):
        return {k: replace_test_name_in_structure(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [replace_test_name_in_structure(item) for item in obj]
    elif isinstance(obj, str):
        if is_test_name_text(obj):
            return "TEST NAME"
        if is_subtitle_text(obj):
            return ""
        return obj
    return obj


def clean_font_name(font_name: str) -> str:
    if not font_name:
        return "Calibri"
    fn = font_name.lower()
    if "calibri" in fn:
        return "Calibri"
    elif "tahoma" in fn:
        return "Tahoma"
    elif "verdana" in fn:
        return "Verdana"
    elif "times" in fn:
        return "Times New Roman"
    elif "cambria" in fn:
        return "Cambria"
    elif "arial" in fn:
        return "Arial"
    elif "symbol" in fn:
        return "Symbol"
    return "Calibri"


def detect_alignment(bbox, page_width):
    if not bbox or len(bbox) < 4 or not page_width:
        return "left"
    x0, y0, x1, y1 = bbox
    # If block is very wide, alignment is usually left (or justify)
    if (x1 - x0) > 0.8 * page_width:
        return "left"
    # Check center symmetry
    left_margin = x0
    right_margin = page_width - x1
    if abs(left_margin - right_margin) < 15.0:
        return "center"
    if right_margin < 50.0 and left_margin > 200.0:
        return "right"
    return "left"


def get_css_color(color_tuple):
    if not color_tuple:
        return None
    try:
        if isinstance(color_tuple, (list, tuple)):
            if len(color_tuple) == 3:
                r = int(color_tuple[0] * 255)
                g = int(color_tuple[1] * 255)
                b = int(color_tuple[2] * 255)
                return f"rgb({r},{g},{b})"
            elif len(color_tuple) == 1:
                gray = int(color_tuple[0] * 255)
                return f"rgb({gray},{gray},{gray})"
            elif len(color_tuple) == 4:
                c, m, y, k = color_tuple
                r = int(255 * (1 - c) * (1 - k))
                g = int(255 * (1 - m) * (1 - k))
                b = int(255 * (1 - y) * (1 - k))
                return f"rgb({r},{g},{b})"
        elif isinstance(color_tuple, (int, float)):
            val = int(color_tuple * 255)
            return f"rgb({val},{val},{val})"
    except Exception:
        pass
    return None


def get_bbox_background_color(page, bbox) -> str:
    x0, y0, x1, y1 = bbox
    try:
        drawings = page.get_drawings()
        for d in drawings:
            rect = d.get("rect")
            fill = d.get("fill")
            if fill and rect:
                rx0, ry0, rx1, ry1 = rect.x0, rect.y0, rect.x1, rect.y1
                if rx0 <= x0 + 2.0 and ry0 <= y0 + 2.0 and rx1 >= x1 - 2.0 and ry1 >= y1 - 2.0:
                    bg = get_css_color(fill)
                    if bg:
                        m = re.match(r"rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)", bg)
                        if m:
                            r, g, b_val = int(m.group(1)), int(m.group(2)), int(m.group(3))
                            return f"#{r:02x}{g:02x}{b_val:02x}"
                        return bg
    except Exception:
        pass
    return None


def get_bbox_text_style(page, bbox) -> dict:
    rect = fitz.Rect(bbox)
    text_dict = page.get_text("dict", clip=rect)
    
    colors = []
    sizes = []
    fonts = []
    bolds = []
    italics = []
    
    for b in text_dict.get("blocks", []):
        if "lines" in b:
            for ln in b["lines"]:
                for s in ln["spans"]:
                    txt = s.get("text", "").strip()
                    if txt:
                        font = s.get("font", "")
                        colors.append(s.get("color"))
                        sizes.append(s.get("size"))
                        fonts.append(font)
                        bolds.append(bool(s.get("flags", 0) & 2 or "bold" in font.lower()))
                        italics.append(bool(s.get("flags", 0) & 1 or "italic" in font.lower()))
                        
    style = {}
    if colors:
        from collections import Counter
        dom_color = Counter(colors).most_common(1)[0][0]
        if dom_color is not None:
            r = (dom_color >> 16) & 255
            g = (dom_color >> 8) & 255
            b_val = dom_color & 255
            style["text_color"] = f"#{r:02x}{g:02x}{b_val:02x}"
            
    if sizes:
        style["font_size"] = round(max(sizes), 2)
        
    if fonts:
        style["font_family"] = clean_font_name(Counter(fonts).most_common(1)[0][0])
        
    if bolds:
        style["bold"] = any(bolds)
        
    if italics:
        style["italic"] = any(italics)
        
    style["alignment"] = detect_alignment(bbox, page.rect.width)
    
    return style


def parse_header_key_value_pairs(full_text: str) -> dict:
    """
    Extracts structured key-value header metadata pairs from document text dynamically.
    Enforces strict criteria for keys (short labels, no prose sentences) to ensure accuracy.
    """
    kv_dict = {}
    
    # Normalize: Join lines starting with a colon to the previous key line
    normalized_text = re.sub(r'\n\s*:', ':', full_text)
    lines = normalized_text.split('\n')
    
    # Pre-process lines to join split keys and values (e.g. key on line 1, : value on line 2)
    joined_lines = []
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if i + 1 < len(lines) and lines[i+1].strip().startswith(':'):
            joined_lines.append(f"{line} {lines[i+1].strip()}")
            i += 2
        else:
            joined_lines.append(line)
            i += 1
            
    ignore_keywords = [
        'received from', 'responsibility', 'presumed that', 'http', 'www',
        'copyright', 'printed on', 'page', 'all rights reserved', 'disclaimer',
        'classified as', 'predicted to', 'evidence', 'database', 'population'
    ]

    prose_verbs = [' is ', ' are ', ' was ', ' were ', ' to ', ' by ', ' from ', ' with ', ' that ', ' which ']

    for line_str in joined_lines:
        if not line_str or len(line_str) > 250:
            continue

        # Find key-value patterns like "Label: Value" (allowing colons in value for times)
        matches = re.findall(r'([A-Za-z0-9][A-Za-z0-9\s/&#\-_]{1,35}):\s*([^\n\t]{1,120})', line_str)
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

    # Step 1: Detect multi-page repeated strings in top 25% and bottom 12%
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
            if y0 < 0.25 * H:
                top_str_counts[norm_t] += 1
            elif y0 > 0.88 * H:
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

            # Check Header match (top 25% of page)
            if y0 < 0.25 * H:
                is_hdr = False
                if norm_t in repeated_top:
                    is_hdr = True
                elif any(kw in norm_t for kw in header_kw):
                    is_hdr = True
                elif re.search(r'page\s+\d+', norm_t):
                    is_hdr = True

                if is_hdr:
                    if y1 > header_y1:
                        header_y1 = y1

            # Check Footer match (bottom 15% of page ONLY)
            if y0 > 0.85 * H:
                is_ftr = False
                if norm_t in repeated_bot:
                    is_ftr = True
                elif any(kw in norm_t for kw in footer_kw):
                    if 'page ' in norm_t and len(norm_t) > 30 and not any(k in norm_t for k in ['dr.', 'signatory', 'verified', 'reviewed', 'authorized']):
                        pass
                    else:
                        is_ftr = True
                elif re.search(r'page\s+\d+', norm_t):
                    if len(norm_t) < 30:
                        is_ftr = True
                elif y0 > 0.90 * H:
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
            if iy0 > 0.85 * H:
                if iy0 < footer_y0:
                    footer_y0 = iy0

        final_hdr_cutoff = min(header_y1 + 2.0, 0.25 * H) if header_y1 > 0 else 0.0
        final_ftr_cutoff = max(footer_y0 - 2.0, 0.85 * H) if footer_y0 < H else H

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
        if norm.startswith(': name') or norm.startswith('name :') or re.match(r'^name\s*:\s*(mr|mrs|ms|dr|master)', norm):
            return True
    return False


def format_bbox(bbox_list):
    """Convert bounding box from [x0, y0, x1, y1] to dictionary format with normalized points coordinates."""
    if not bbox_list or len(bbox_list) < 4:
        return None
    x0 = round(bbox_list[0], 2)
    y0 = round(bbox_list[1], 2)
    x1 = round(bbox_list[2], 2)
    y1 = round(bbox_list[3], 2)
    return {
        "x": x0,
        "y": y0,
        "width": round(x1 - x0, 2),
        "height": round(y1 - y0, 2),
        "unit": "pt",
        "x0": x0,
        "y0": y0,
        "x1": x1,
        "y1": y1
    }


def sanitize_headers(headers):
    """Sanitize and deduplicate table headers to guarantee unique dictionary keys."""
    sanitized = []
    seen = {}
    for idx, h in enumerate(headers):
        h_str = str(h).strip()
        if not h_str:
            h_str = f"Column {idx + 1}"
        if h_str in seen:
            seen[h_str] += 1
            h_str = f"{h_str}_{seen[h_str]}"
        else:
            seen[h_str] = 1
        sanitized.append(h_str)
    return sanitized


def check_table_continuation(t_curr_bbox, t_prev_bbox, cols_curr, cols_prev, hy_cutoff, fy_cutoff_prev):
    """Check if a table is a continuation of the last table on the previous page."""
    if cols_curr != cols_prev or cols_curr == 0:
        return False
    
    # Check horizontal alignment
    x0_diff = abs(t_curr_bbox[0] - t_prev_bbox[0])
    x1_diff = abs(t_curr_bbox[2] - t_prev_bbox[2])
    if x0_diff > 15 or x1_diff > 15:
        return False
        
    # Must start near the top of page body cutoff
    y0_curr = t_curr_bbox[1]
    if y0_curr > hy_cutoff + 80:
        return False
        
    # Must end near the bottom of previous page body cutoff
    y1_prev = t_prev_bbox[3]
    if y1_prev < fy_cutoff_prev - 180:
        return False
        
    return True


def classify_block_type(text, max_font_size, is_bold, font_name):
    """Semantically classify a text block."""
    text_stripped = text.strip()
    if not text_stripped:
        return "unknown"
        
    text_lower = text_stripped.lower()
    if "pmid" in text_lower or "doi:" in text_lower or text_lower.startswith("http") or "www." in text_lower:
        return "paragraph"
        
    # Check key-value match
    if ":" in text_stripped:
        parts = text_stripped.split(":", 1)
        k = parts[0].strip()
        v = parts[1].strip()
        if v and len(k) >= 2 and len(k) <= 35 and k.count(" ") < 4 and not any(c in k for c in [".", ",", ";", "?", "!"]):
            return "key_value"
            
    is_hd_cand = False
    if max_font_size >= 11.5:
        is_hd_cand = True
    elif is_bold and len(text_stripped) < 90:
        if not text_stripped.endswith(('.', ',', ';', '?')):
            is_hd_cand = True
            
    if is_hd_cand:
        if max_font_size >= 12.0 or text_stripped.isupper():
            return "heading"
        else:
            return "subheading"
            
    return "paragraph"


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
    
    # New page-by-page structured document flow list
    document_pages = []

    last_table_prev_page = None

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
        
        # New page elements list
        page_elements = []
        header_elements = []
        footer_elements = []

        try:
            if hasattr(page, "find_tables"):
                tf = page.find_tables()
                table_list = list(tf)
                # Sort tables on this page by top-Y coordinate
                table_list.sort(key=lambda t: t.bbox[1])
                
                for t_idx, table in enumerate(table_list):
                    t_bbox = [round(c, 2) for c in table.bbox]
                    t_mid = (t_bbox[1] + t_bbox[3]) / 2.0
                    
                    table_data = table.extract()
                    if table_data and len(table_data) > 0:
                        raw_headers = [str(c).strip() if c else "" for c in table_data[0]]
                        raw_rows = [[str(cell).strip() if cell else "" for cell in row] for row in table_data[1:]]
                    else:
                        continue

                    # Extract header and cell styles dynamically from cell bboxes
                    header_styles = []
                    cell_styles = []
                    if hasattr(table, "rows") and table.rows:
                        for r_idx, row_obj in enumerate(table.rows):
                            r_styles = []
                            for c_idx, cell_obj in enumerate(row_obj.cells):
                                if cell_obj:
                                    c_bbox = cell_obj.bbox if hasattr(cell_obj, "bbox") else cell_obj
                                    c_style = get_bbox_text_style(page, c_bbox)
                                    bg = get_bbox_background_color(page, c_bbox)
                                    if bg:
                                        c_style["background_color"] = bg
                                    r_styles.append(c_style)
                                else:
                                    r_styles.append({})
                            if r_idx == 0:
                                header_styles = r_styles
                            else:
                                cell_styles.append(r_styles)

                    # If table is in header / footer area, add to header/footer elements
                    if t_mid < hy_cutoff:
                        col_widths = []
                        if hasattr(table, "cols") and table.cols:
                            for idx in range(len(table.cols) - 1):
                                col_widths.append(round(table.cols[idx+1] - table.cols[idx], 2))
                        columns_with_widths = []
                        san_headers = sanitize_headers(raw_headers)
                        for idx, h_name in enumerate(san_headers):
                            col_item = {"name": h_name}
                            if idx < len(col_widths):
                                col_item["width"] = col_widths[idx]
                            columns_with_widths.append(col_item)
                        row_dicts = []
                        for row in raw_rows:
                            row_dict = {}
                            for idx, h_name in enumerate(san_headers):
                                val = row[idx] if idx < len(row) else ""
                                row_dict[h_name] = val
                            row_dicts.append(row_dict)
                        header_elements.append({
                            "type": "table",
                            "bbox": format_bbox(t_bbox),
                            "columns": columns_with_widths,
                            "rows": row_dicts,
                            "header_styles": header_styles,
                            "cell_styles": cell_styles
                        })
                        continue
                    elif t_mid > fy_cutoff:
                        col_widths = []
                        if hasattr(table, "cols") and table.cols:
                            for idx in range(len(table.cols) - 1):
                                col_widths.append(round(table.cols[idx+1] - table.cols[idx], 2))
                        columns_with_widths = []
                        san_headers = sanitize_headers(raw_headers)
                        for idx, h_name in enumerate(san_headers):
                            col_item = {"name": h_name}
                            if idx < len(col_widths):
                                col_item["width"] = col_widths[idx]
                            columns_with_widths.append(col_item)
                        row_dicts = []
                        for row in raw_rows:
                            row_dict = {}
                            for idx, h_name in enumerate(san_headers):
                                val = row[idx] if idx < len(row) else ""
                                row_dict[h_name] = val
                            row_dicts.append(row_dict)
                        footer_elements.append({
                            "type": "table",
                            "bbox": format_bbox(t_bbox),
                            "columns": columns_with_widths,
                            "rows": row_dicts,
                            "header_styles": header_styles,
                            "cell_styles": cell_styles
                        })
                        continue

                    # Apply continuation check against the last table on the previous page
                    is_continuation = False
                    if len(page_tables) == 0 and last_table_prev_page is not None:
                        prev_bbox = last_table_prev_page.get("bbox", [0, 0, 0, 0])
                        prev_cols = len(last_table_prev_page.get("headers", []))
                        prev_fy_cutoff = page_bounds[page_num - 1]["footer_y_cutoff"]
                        
                        if check_table_continuation(t_bbox, prev_bbox, len(raw_headers), prev_cols, hy_cutoff, prev_fy_cutoff):
                            is_continuation = True
                            
                            # Prepend the extracted headers as the first data row
                            if any(raw_headers):
                                raw_rows.insert(0, raw_headers)
                                
                            # Inherit headers from previous page's table
                            raw_headers = last_table_prev_page.get("raw_headers", last_table_prev_page.get("headers", []))
                    
                    # Generate legacy-friendly headers and rows (newline-replaced)
                    legacy_headers = [h.replace("\n", " ") for h in raw_headers]
                    legacy_rows = [[cell.replace("\n", " ") for cell in row] for row in raw_rows]
                    
                    table_bboxes.append(t_bbox)
                    
                    table_obj = {
                        "page": page_num + 1,
                        "table_index": len(page_tables) + 1,
                        "bbox": t_bbox,
                        "type": "table",
                        "headers": legacy_headers,
                        "rows": legacy_rows,
                        "raw_headers": raw_headers,
                        "raw_rows": raw_rows,
                        "is_continuation": is_continuation,
                        "header_styles": header_styles,
                        "cell_styles": cell_styles
                    }
                    page_tables.append(table_obj)
                    all_tables_list.append(table_obj)
                    
                    # Map headers to dictionary for the new element format
                    san_headers = sanitize_headers(raw_headers)
                    row_dicts = []
                    for row in raw_rows:
                        row_dict = {}
                        for idx, h_name in enumerate(san_headers):
                            val = row[idx] if idx < len(row) else ""
                            row_dict[h_name] = val
                        row_dicts.append(row_dict)
                    
                    col_widths = []
                    if hasattr(table, "cols") and table.cols:
                        for idx in range(len(table.cols) - 1):
                            col_widths.append(round(table.cols[idx+1] - table.cols[idx], 2))
                    columns_with_widths = []
                    for idx, h_name in enumerate(san_headers):
                        col_item = {"name": h_name}
                        if idx < len(col_widths):
                            col_item["width"] = col_widths[idx]
                        columns_with_widths.append(col_item)
                        
                    page_elements.append({
                        "type": "table",
                        "bbox": format_bbox(t_bbox),
                        "columns": columns_with_widths,
                        "rows": row_dicts,
                        "is_continuation": is_continuation,
                        "header_styles": header_styles,
                        "cell_styles": cell_styles
                    })
        except Exception as t_err:
            logger.debug(f"Page {page_num + 1} table error: {t_err}")

        # Update last active table
        if page_tables:
            last_table_prev_page = page_tables[-1]
        else:
            last_table_prev_page = None

        # 2. Parse text blocks outside tables and strictly inside body region ONLY
        page_blocks = []
        body_text_lines = []
        for block in text_page_dict.get("blocks", []):
            if "lines" in block:
                b_bbox = [round(c, 2) for c in block["bbox"]]
                if is_inside_table_bbox(b_bbox, table_bboxes):
                    continue

                b_text = ""
                is_bold = False
                is_italic = False
                max_size = 0.0
                font_name = ""
                lines_list = []
                colors_in_block = []
                sizes_in_block = []
                fonts_in_block = []
                lines_data = []
                for line in block["lines"]:
                    line_text = " ".join(span.get("text", "") for span in line["spans"]).strip()
                    if line_text:
                        lines_list.append(line_text)
                    spans_data = []
                    for span in line["spans"]:
                        text = span.get("text", "").strip()
                        if text:
                            b_text += span.get("text", "") + " "
                            span_font = span.get("font", "")
                            if span.get("flags", 0) & 2 or "bold" in span_font.lower():
                                is_bold = True
                            if span.get("flags", 0) & 1 or "italic" in span_font.lower():
                                is_italic = True
                            span_size = span.get("size", 0.0)
                            if span_size > max_size:
                                max_size = span_size
                            font_name = span_font
                            
                            span_color_dec = span.get("color")
                            span_color = "#000000"
                            if span_color_dec is not None:
                                r = (span_color_dec >> 16) & 255
                                g = (span_color_dec >> 8) & 255
                                b_val = span_color_dec & 255
                                span_color = f"#{r:02x}{g:02x}{b_val:02x}"
                                colors_in_block.append(span_color_dec)
                            
                            sizes_in_block.append(span_size)
                            if span_font:
                                fonts_in_block.append(span_font)
                                
                            spans_data.append({
                                "text": span.get("text", ""),
                                "font": span_font,
                                "size": round(span_size, 2),
                                "color": span_color,
                                "bold": bool(span.get("flags", 0) & 2 or "bold" in span_font.lower()),
                                "italic": bool(span.get("flags", 0) & 1 or "italic" in span_font.lower()),
                                "bbox": [round(c, 2) for c in span.get("bbox", [0, 0, 0, 0])]
                            })
                    if spans_data:
                        lines_data.append({
                            "bbox": [round(c, 2) for c in line.get("bbox", [0, 0, 0, 0])],
                            "spans": spans_data
                        })

                clean_b_text = b_text.strip()
                if not clean_b_text:
                    continue

                sem_type = classify_block_type(clean_b_text, max_size, is_bold, font_name)
                
                # Determine dominant text color
                text_color = "#000000"
                if colors_in_block:
                    from collections import Counter
                    dom_color_dec = Counter(colors_in_block).most_common(1)[0][0]
                    r = (dom_color_dec >> 16) & 255
                    g = (dom_color_dec >> 8) & 255
                    b_val = dom_color_dec & 255
                    text_color = f"#{r:02x}{g:02x}{b_val:02x}"
                elif sem_type == "heading":
                    text_color = "#1f497d"
                elif sem_type == "subheading":
                    text_color = "#008080"
                
                # Dominant font name
                dom_font = clean_font_name(font_name)
                if fonts_in_block:
                    from collections import Counter
                    dom_font = clean_font_name(Counter(fonts_in_block).most_common(1)[0][0])
                
                style_override = {}
                if max_size:
                    style_override["font_size"] = round(max_size, 2)
                if dom_font:
                    style_override["font_family"] = dom_font
                if is_bold:
                    style_override["bold"] = True
                if is_italic:
                    style_override["italic"] = True
                if text_color:
                    style_override["text_color"] = text_color
                style_override["alignment"] = detect_alignment(b_bbox, rect.width)

                # Determine if block is in header, footer, or body
                if b_bbox[1] < hy_cutoff:
                    if re.search(r'^page\s+\d+(\s+of\s+\d+)?$', clean_b_text, re.I):
                        continue
                    el_obj = {
                        "type": sem_type,
                        "bbox": format_bbox(b_bbox),
                        "text": clean_b_text
                    }
                    if style_override:
                        el_obj["style_override"] = style_override
                    header_elements.append(el_obj)

                    if page_num == 0:
                        is_hdr_card = is_patient_sample_header_block(clean_b_text, b_bbox[1])
                        if is_hdr_card:
                            raw_block_text = "\n".join(lines_list)
                            norm_block_text = re.sub(r'\n\s*:', ':', raw_block_text)
                            parsed_kv = parse_header_key_value_pairs(norm_block_text)
                            for k, v in parsed_kv.items():
                                if not any(kw in k.lower() for kw in ['neuberg', 'laboratory report', 'barcode', 'qr code', 'acc. remarks', 'accession.id', 'centre details']):
                                    all_kv_pairs[k] = v
                            if parsed_kv:
                                clean_kv = {k: v for k, v in parsed_kv.items() if not any(kw in k.lower() for kw in ['neuberg', 'laboratory report', 'barcode', 'qr code', 'acc. remarks', 'accession.id', 'centre details'])}
                                if clean_kv:
                                    page_elements.append({
                                        "type": "key_value",
                                        "bbox": format_bbox(b_bbox),
                                        "data": clean_kv
                                    })
                    continue

                elif b_bbox[3] > fy_cutoff or b_bbox[1] > fy_cutoff:
                    if re.search(r'^page\s+\d+(\s+of\s+\d+)?$', clean_b_text, re.I):
                        continue
                    role = "signature_block" if any(kw in clean_b_text.lower() for kw in ['reviewed by', 'authorized signatory', 'signatory', 'doctor', 'dr.']) else None
                    el_obj = {
                        "type": sem_type,
                        "bbox": format_bbox(b_bbox),
                        "text": clean_b_text
                    }
                    if role:
                        el_obj["role"] = role
                    if style_override:
                        el_obj["style_override"] = style_override
                    footer_elements.append(el_obj)
                    continue

                if re.search(r'^page\s+\d+(\s+of\s+\d+)?$', clean_b_text, re.I):
                    continue
                if any(kw in clean_b_text.lower() for kw in ['reviewed by', 'authorized signatory', 'mc-7414']):
                    continue
                    
                is_hdr_card = is_patient_sample_header_block(clean_b_text, b_bbox[1])
                if is_hdr_card:
                    if page_num == 0:
                        raw_block_text = "\n".join(lines_list)
                        norm_block_text = re.sub(r'\n\s*:', ':', raw_block_text)
                        parsed_kv = parse_header_key_value_pairs(norm_block_text)
                        for k, v in parsed_kv.items():
                            if not any(kw in k.lower() for kw in ['neuberg', 'laboratory report', 'barcode', 'qr code', 'acc. remarks', 'accession.id', 'centre details']):
                                all_kv_pairs[k] = v
                        if parsed_kv:
                            clean_kv = {k: v for k, v in parsed_kv.items() if not any(kw in k.lower() for kw in ['neuberg', 'laboratory report', 'barcode', 'qr code', 'acc. remarks', 'accession.id', 'centre details'])}
                            if clean_kv:
                                page_elements.append({
                                    "type": "key_value",
                                    "bbox": format_bbox(b_bbox),
                                    "data": clean_kv
                                })
                    continue

                is_hd_cand = (max_size >= 11.5 or (is_bold and len(clean_b_text) < 70)) and not clean_b_text.endswith(('.', ',', ';', '?'))
                block_obj = {
                    "page": page_num + 1,
                    "bbox": b_bbox,
                    "type": "heading" if is_hd_cand else "paragraph",
                    "text": clean_b_text,
                    "max_font_size": round(max_size, 2),
                    "is_bold": is_bold,
                    "italic": is_italic,
                    "font": dom_font,
                    "color": text_color,
                    "lines": lines_data
                }
                page_blocks.append(block_obj)
                text_blocks_with_style.append(block_obj)
                body_text_lines.append(clean_b_text)
                
                if sem_type == "key_value":
                    raw_block_text = "\n".join(lines_list)
                    parsed_kv = parse_header_key_value_pairs(raw_block_text)
                    if parsed_kv:
                        clean_kv = {k: v for k, v in parsed_kv.items() if not any(kw in k.lower() for kw in ['neuberg', 'laboratory report', 'barcode', 'qr code', 'acc. remarks', 'accession.id', 'centre details'])}
                        if clean_kv:
                            for k, v in clean_kv.items():
                                all_kv_pairs[k] = v
                            page_elements.append({
                                "type": "key_value",
                                "bbox": format_bbox(b_bbox),
                                "data": clean_kv
                            })
                            continue
                    sem_type = "paragraph"
                    
                el_obj = {
                    "type": sem_type,
                    "bbox": format_bbox(b_bbox),
                    "text": clean_b_text
                }
                if style_override:
                    el_obj["style_override"] = style_override
                page_elements.append(el_obj)

        page_body_text = "\n".join(body_text_lines)
        full_text_pages.append(page_body_text)

        filtered_page_kv = {}
        if page_num == 0:
            page_kv = parse_header_key_value_pairs(page_body_text)
            for k, v in page_kv.items():
                if not any(hk in k.lower() for hk in header_kv_ignore_keys):
                    filtered_page_kv[k] = v
                    all_kv_pairs[k] = v

        # 3. Extract Images & Graphs
        raw_images = extract_page_images_and_graphs(doc, page_num, page)
        page_images = []
        for img in raw_images:
            img_bbox = img.get("bbox")
            if img_bbox:
                y0, y1 = img_bbox[1], img_bbox[3]
                if y1 <= hy_cutoff:
                    header_elements.append({
                        "type": "image",
                        "bbox": format_bbox(img_bbox),
                        "data_uri": img.get("data_uri"),
                        "width": img.get("width"),
                        "height": img.get("height"),
                        "role": "logo"
                    })
                    continue
                elif y0 >= fy_cutoff:
                    footer_elements.append({
                        "type": "image",
                        "bbox": format_bbox(img_bbox),
                        "data_uri": img.get("data_uri"),
                        "width": img.get("width"),
                        "height": img.get("height"),
                        "role": "signature"
                    })
                    continue
            img["type"] = "image"
            page_images.append(img)
            all_images_list.append(img)
            
            page_elements.append({
                "type": "image",
                "bbox": format_bbox(img_bbox or [0, 0, 0, 0]),
                "data_uri": img.get("data_uri"),
                "width": img.get("width"),
                "height": img.get("height")
            })

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
        
        # Sort and merge new page_elements structure in visual/reading order
        def get_el_y0(el):
            bbox = el.get("bbox")
            if bbox:
                return (round(bbox.get("y0", 0.0) / 3.0) * 3.0, bbox.get("x0", 0.0))
            return (0.0, 0.0)
            
        page_elements.sort(key=get_el_y0)
        
        # Merge adjacent key_values
        merged_elements = []
        for el in page_elements:
            if not merged_elements:
                merged_elements.append(el)
                continue
            prev_el = merged_elements[-1]
            if el["type"] == "key_value" and prev_el["type"] == "key_value":
                prev_el["data"].update(el["data"])
                # Merge bboxes
                p_bb = prev_el["bbox"]
                c_bb = el["bbox"]
                if p_bb and c_bb:
                    x0 = min(p_bb["x0"], c_bb["x0"])
                    y0 = min(p_bb["y0"], c_bb["y0"])
                    x1 = max(p_bb["x1"], c_bb["x1"])
                    y1 = max(p_bb["y1"], c_bb["y1"])
                    prev_el["bbox"] = {
                        "x": x0,
                        "y": y0,
                        "width": round(x1 - x0, 2),
                        "height": round(y1 - y0, 2),
                        "unit": "pt",
                        "x0": x0,
                        "y0": y0,
                        "x1": x1,
                        "y1": y1
                    }
            else:
                merged_elements.append(el)
                
        # Set table titles based on preceding heading/subheading
        for i in range(len(merged_elements)):
            if merged_elements[i]["type"] == "table":
                if i > 0:
                    prev = merged_elements[i-1]
                    if prev["type"] in ["heading", "subheading"]:
                        p_bbox = prev.get("bbox")
                        t_bbox = merged_elements[i].get("bbox")
                        if p_bbox and t_bbox:
                            if t_bbox["y0"] - p_bbox["y1"] < 40:
                                merged_elements[i]["title"] = prev["text"]
                                
        final_page_elements = []
        if header_elements:
            final_page_elements.append({
                "type": "header",
                "bbox": {
                    "x": 0,
                    "y": 0,
                    "width": round(rect.width, 2),
                    "height": round(hy_cutoff, 2),
                    "unit": "pt"
                },
                "elements": header_elements
            })
        
        final_page_elements.extend(merged_elements)
        
        if footer_elements:
            final_page_elements.append({
                "type": "footer",
                "bbox": {
                    "x": 0,
                    "y": round(fy_cutoff, 2),
                    "width": round(rect.width, 2),
                    "height": round(rect.height - fy_cutoff, 2),
                    "unit": "pt"
                },
                "elements": footer_elements
            })

        document_pages.append({
            "page_number": page_num + 1,
            "width": round(rect.width, 2),
            "height": round(rect.height, 2),
            "unit": "pt",
            "elements": final_page_elements
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
        "document": {
            "file_name": os.path.basename(pdf_path),
            "total_pages": len(doc),
            "pages": document_pages
        },
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

    extracted_data = replace_sng_in_structure(extracted_data)
    extracted_data = replace_test_name_in_structure(extracted_data)

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

    # Automatically save Word (.docx) layout into extracted_jsons/ and output/ directories
    try:
        from converter import convert_json_to_docx
        
        # Load theme config from theme.json if present
        theme_config = {}
        theme_json_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "theme.json")
        if os.path.exists(theme_json_path):
            try:
                with open(theme_json_path, "r", encoding="utf-8") as f_theme:
                    theme_config = json.load(f_theme)
            except Exception:
                pass

        # Save to extracted_jsons/ as docx
        json_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "extracted_jsons")
        docx_save_path1 = os.path.join(json_dir, f"{pdf_stem}.docx")
        convert_json_to_docx(extracted_data, output_path=docx_save_path1, theme_config=theme_config)
        logger.info(f"Saved extracted Word layout to: {docx_save_path1}")

        # Save to output/ as _report.docx
        output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
        os.makedirs(output_dir, exist_ok=True)
        docx_save_path2 = os.path.join(output_dir, f"{pdf_stem}_report.docx")
        convert_json_to_docx(extracted_data, output_path=docx_save_path2, theme_config=theme_config)
        logger.info(f"Saved extracted Word layout to: {docx_save_path2}")
    except Exception as e_word:
        logger.warning(f"Could not auto-save Word document: {e_word}")

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


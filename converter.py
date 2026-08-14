"""
Universal Dynamic PDF HTML Renderer & Converter Module
======================================================
Provides rendering modes and export tools:
1. Exact Position Layout Mode (`render_exact_pdf_layout_html`): Preserves 100% exact 1-to-1 visual
   positions (`top`, `left`, `width`, `height`) of all text, tables, images, banners, and boxes from input PDF.
2. Standard Flow Template Mode (`generate_dynamic_template_html`): Renders extracted JSON into a responsive web flow.
3. Microsoft Word Exporter (`convert_json_to_docx`): Generates formatted Word .docx files from extracted JSON.
"""

import os
import io
import json
import base64
import re
from pathlib import Path
from PIL import Image
import fitz  # PyMuPDF

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    sync_playwright = None

try:
    from pdf2docx import Converter as PDF2DocxConverter
except ImportError:
    PDF2DocxConverter = None

from extractor import extract_report_data, detect_dynamic_header_footer_bounds


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
    return "black"


def render_exact_pdf_layout_html(doc, doc_title: str = "Uploaded Document", theme_config: dict = None) -> str:
    """
    Renders an HTML document where body elements stay in their exact visual positions,
    while original header and footer content is completely excluded dynamically.
    Applies user-selected theme typography, primary colors, table header styling, and cell borders.
    """
    cfg = get_merged_theme_config(theme_config)
    colors_cfg = cfg.get("colors", {})
    typo_cfg = cfg.get("typography", {})

    primary_color = colors_cfg.get("primary", "#1f497d")
    secondary_color = colors_cfg.get("secondary", "#008080")
    accent_orange = colors_cfg.get("accent_orange", "#ed7d31")
    accent_red = colors_cfg.get("accent_red", "#ff0000")
    banner_dark = colors_cfg.get("banner_dark", "#404040")
    text_primary = colors_cfg.get("text_primary", "#000000")
    text_dark = colors_cfg.get("text_dark", "#0d0d0d")
    bg_page = colors_cfg.get("background_page", "#ffffff")
    border_color = colors_cfg.get("border_primary", primary_color)
    font_family = typo_cfg.get("primary_family", "Cambria, 'Times New Roman', serif")

    fallback_left = 35.5
    fallback_width = 524.0

    # Detect dynamic header and footer bounds per page
    page_bounds = detect_dynamic_header_footer_bounds(doc)

    html_parts = [
        "<!DOCTYPE html>",
        "<html>",
        "<head>",
        "<meta charset='utf-8'>",
        f"<title>{doc_title}</title>",
        "<style>",
        f"""
        :root {{
          --color-primary: {primary_color};
          --color-secondary: {secondary_color};
          --color-accent-orange: {accent_orange};
          --color-accent-red: {accent_red};
          --color-banner-dark: {banner_dark};
          --color-bg-page: {bg_page};
          --color-bg-container: #525659;
          --text-primary: {text_primary};
          --text-teal: {secondary_color};
          --text-orange: {accent_orange};
          --text-red: {accent_red};
          --border-color: {border_color};
          --font-primary: {font_family};
        }}
        @page {{ size: 595.6pt 842.0pt; margin: 0; }}
        * {{ box-sizing: border-box; }}
        body {{ margin: 0; padding: 0; background-color: var(--color-bg-container); font-family: var(--font-primary); color: var(--text-primary); }}
        .pdf-container {{ display: flex; flex-direction: column; align-items: center; padding: 20px 0; }}
        .pdf-page {{ background: var(--color-bg-page); width: 595.6pt; min-height: 842.0pt; margin-bottom: 20px; position: relative; overflow: visible; box-shadow: 0 4px 12px rgba(0,0,0,0.3); page-break-after: always; font-family: var(--font-primary); }}
        div[id^='page'] {{ position: relative !important; width: 595.6pt !important; min-height: 842.0pt !important; overflow: visible !important; }}
        div[id^='page'] p {{ position: absolute !important; margin: 0 !important; padding: 0 !important; white-space: normal !important; word-break: normal !important; overflow-wrap: break-word !important; max-width: 100% !important; z-index: 10 !important; font-family: var(--font-primary) !important; line-height: 1.2 !important; overflow: visible !important; }}
        div[id^='page'] span {{ word-break: normal !important; overflow-wrap: break-word !important; }}
        .black-banner-span {{ color: #ffffff !important; display: block !important; width: 100% !important; text-align: center !important; padding: 4px 0 !important; line-height: 1.2 !important; font-weight: bold !important; font-size: 13.0pt !important; font-family: var(--font-primary) !important; border-radius: 0px !important; white-space: normal !important; margin: 0 !important; background-color: var(--color-banner-dark) !important; box-sizing: border-box !important; z-index: 15 !important; }}
        .label-bar-span {{ color: #ffffff !important; display: inline-block !important; padding: 2px 6px !important; line-height: 1.2 !important; border-radius: 2px !important; font-family: var(--font-primary) !important; word-break: break-word !important; margin: 0 !important; background-color: var(--color-primary) !important; z-index: 15 !important; }}
        .table-header-cell {{ position: absolute; background-color: var(--color-primary) !important; color: #ffffff !important; font-family: var(--font-primary) !important; font-size: 9.5pt !important; font-weight: bold !important; display: flex !important; align-items: center !important; justify-content: center !important; text-align: center !important; padding: 2px 4px !important; white-space: normal !important; word-break: break-word !important; overflow-wrap: break-word !important; line-height: 1.15 !important; border: 1px solid var(--border-color) !important; box-sizing: border-box !important; z-index: 15 !important; pointer-events: none !important; }}
        div[id^='page'] img {{ position: absolute !important; transform-origin: 0 0 !important; z-index: 5 !important; opacity: 1 !important; visibility: visible !important; display: inline-block !important; }}
        .table-grid-cell {{ position: absolute; border: 1px solid var(--border-color) !important; background: transparent; pointer-events: none; z-index: 2 !important; }}
        .vector-box {{ position: absolute; border: 1px solid var(--border-color) !important; background: transparent; pointer-events: none; z-index: 2 !important; border-radius: 2px; box-sizing: border-box !important; }}
        .vector-fill-box {{ position: absolute; background-color: var(--color-primary) !important; pointer-events: none; z-index: 2 !important; border-radius: 2px; box-sizing: border-box !important; }}
        .section-content-box {{ position: absolute; border: 1px solid var(--border-color) !important; background: transparent; pointer-events: none; z-index: 2 !important; border-radius: 3px; }}
        .teal-text {{ color: var(--color-secondary) !important; font-weight: bold; }}
        .orange-text {{ color: var(--color-accent-orange) !important; font-weight: bold; }}
        .red-text {{ color: var(--color-accent-red) !important; font-weight: bold; }}
        @media print {{ body {{ background-color: #ffffff; }} .pdf-container {{ padding: 0; }} .pdf-page {{ margin: 0; box-shadow: none; }} }}
        """,
        "</style>",
        "</head>",
        "<body>",
        "<div class='pdf-container'>",
    ]

    for page_num in range(len(doc)):
        page = doc[page_num]
        page_left_val, page_width_val = _get_page_bounds(page, fallback_left, fallback_width)
        page_left_str = f"{page_left_val:.1f}pt"
        page_width_str = f"{page_width_val:.1f}pt"

        hy_cutoff = page_bounds[page_num]["header_y_cutoff"]
        fy_cutoff = page_bounds[page_num]["footer_y_cutoff"]

        html_parts.append(f"<div class='pdf-page' id='page-{page_num+1}'>")
        page_html = page.get_text("html")

        # 1. Extract PyMuPDF table headers & grid coordinates in body region ONLY
        tabs = page.find_tables()
        table_header_html_divs = []
        table_grid_html_divs = []
        header_y_ranges = []

        for tab in tabs.tables:
            if hasattr(tab, 'bbox'):
                tx0, ty0, tx1, ty1 = tab.bbox
                if ty0 < hy_cutoff or ty1 > fy_cutoff:
                    continue

            valid_cells = [c for c in tab.cells if c]
            if valid_cells:
                min_y0 = min(c[1] for c in valid_cells)
                header_cells = [c for c in valid_cells if abs(c[1] - min_y0) < 3.0]
                header_cells.sort(key=lambda c: c[0])
                if len(header_cells) >= 2:
                    hy0 = min(c[1] for c in header_cells)
                    hy1 = max(c[3] for c in header_cells)
                    header_y_ranges.append((hy0 - 3.0, hy1 + 3.0))

                    for c in header_cells:
                        x0, y0, x1, y1 = c
                        w = x1 - x0
                        h = max(24.0, y1 - y0)
                        rect = fitz.Rect(x0, y0, x1, y1)
                        raw_text = page.get_text('text', clip=rect).strip()
                        formatted_text = raw_text.replace('\n', ' ').strip()
                        if "/" in formatted_text and " " not in formatted_text:
                            formatted_text = formatted_text.replace("/", "/ ")
                        table_header_html_divs.append(
                            f"<div class='table-header-cell' "
                            f"style='left:{x0:.1f}pt;top:{y0:.1f}pt;"
                            f"width:{w:.1f}pt;height:{h:.1f}pt;'>"
                            f"{formatted_text}</div>"
                        )

            for cell in tab.cells:
                if cell:
                    cx0, cy0, cx1, cy1 = cell
                    cw = cx1 - cx0
                    ch = cy1 - cy0
                    if any(hy0_r <= cy0 <= hy1_r for hy0_r, hy1_r in header_y_ranges):
                        continue
                    table_grid_html_divs.append(
                        f"<div class='table-grid-cell' "
                        f"style='left:{cx0:.1f}pt;top:{cy0:.1f}pt;"
                        f"width:{cw:.1f}pt;height:{ch:.1f}pt;'></div>"
                    )

        # 2. Extract Vector Drawings in body region ONLY
        vector_html_divs = []
        try:
            drawings = page.get_drawings()
            for d in drawings:
                rect = d.get('rect')
                if not rect:
                    continue
                rx0, ry0, rx1, ry1 = rect.x0, rect.y0, rect.x1, rect.y1
                rw, rh = rx1 - rx0, ry1 - ry0
                
                # Skip tiny points and full page borders
                if rw < 1.0 and rh < 1.0:
                    continue
                if rw > 550 and rh > 800:
                    continue

                # Exclude vector lines/drawings in header or footer regions
                if ry0 < hy_cutoff or ry1 > fy_cutoff:
                    continue

                if any(abs(ry0 - hy0_r) < 5.0 for hy0_r, _ in header_y_ranges):
                    continue

                fill_col = get_css_color(d.get('fill'))
                stroke_col = get_css_color(d.get('color'))
                stroke_w = d.get('width') or 1.0

                if rh <= 1.5:  # Horizontal line segment
                    bg_color = stroke_col or "black"
                    vector_html_divs.append(
                        f"<div style='position:absolute; left:{rx0:.1f}pt; top:{ry0:.1f}pt; width:{rw:.1f}pt; height:{stroke_w:.1f}pt; background-color:{bg_color}; z-index:2; pointer-events:none;'></div>"
                    )
                elif rw <= 1.5:  # Vertical line segment
                    bg_color = stroke_col or "black"
                    vector_html_divs.append(
                        f"<div style='position:absolute; left:{rx0:.1f}pt; top:{ry0:.1f}pt; width:{stroke_w:.1f}pt; height:{rh:.1f}pt; background-color:{bg_color}; z-index:2; pointer-events:none;'></div>"
                    )
                else:  # Rectangle box
                    style_parts = [
                        "position:absolute;",
                        f"left:{rx0:.1f}pt;",
                        f"top:{ry0:.1f}pt;",
                        f"width:{rw:.1f}pt;",
                        f"height:{rh:.1f}pt;",
                        "pointer-events:none;"
                    ]
                    if fill_col:
                        # Use exact background fill color
                        style_parts.append(f"background-color:{fill_col};")
                        style_parts.append("z-index:1;")
                    else:
                        style_parts.append("background-color:transparent;")
                        style_parts.append("z-index:2;")

                    if stroke_col:
                        style_parts.append(f"border:{stroke_w:.1f}pt solid {stroke_col};")

                    vector_html_divs.append(
                        f"<div style='{' '.join(style_parts)}'></div>"
                    )
        except Exception:
            pass

        # 3. Clean raw HTML & suppress raw <p> tags inside header/footer regions or table headers
        cleaned = page_html
        cleaned = re.sub(r'<img\s+[^>]*>', '', cleaned)
        cleaned = re.sub(r'font-family:[^;"]+', f'font-family: {font_family}', cleaned)

        # HIDE raw <p> tags that fall inside header/footer regions or table headers
        def filter_hdr_ftr_and_table_p(match):
            p_tag = match.group(0)
            top_m = re.search(r'top:\s*([\d.]+)pt', p_tag)
            if top_m:
                y_val = float(top_m.group(1))
                if y_val < hy_cutoff or y_val > fy_cutoff:
                    return ""
                if any(hy0_r <= y_val <= hy1_r for hy0_r, hy1_r in header_y_ranges):
                    return re.sub(r'>([^<]+)<', '><', p_tag)
            return p_tag

        cleaned = re.sub(r'<p\s+[^>]*>.*?</p>', filter_hdr_ftr_and_table_p, cleaned, flags=re.DOTALL)

        # 4. Format Black Banners & Blue Section Label Bars
        def format_heading_p(match):
            p_tag = match.group(0)
            is_white_text = bool(re.search(r'color:\s*(?:#ffffff|#fff|rgb\(\s*255\s*,\s*255\s*,\s*255\s*\))', p_tag, re.IGNORECASE))
            size_m = re.search(r'font-size:\s*([\d.]+)pt', p_tag)
            font_sz = float(size_m.group(1)) if size_m else 10.0
            
            top_m = re.search(r'top:\s*([\d.]+)pt', p_tag)
            top_val = top_m.group(1) if top_m else "100.0"

            if is_white_text and font_sz >= 12.0:
                text_val = re.sub(r'<[^>]+>', '', p_tag).strip()
                if len(text_val) > 15:
                    return (
                        f'<p style="top:{top_val}pt;left:{page_left_str};'
                        f'width:{page_width_str};margin:0;padding:0;z-index:15;">'
                        f'<span class="black-banner-span" '
                        f'style="width:{page_width_str};background-color:#404040;">'
                        f'{text_val}</span></p>'
                    )
            return p_tag

        cleaned = re.sub(r'<p\s+[^>]*>.*?</p>', format_heading_p, cleaned, flags=re.DOTALL)

        def format_labelbar_p(match):
            p_tag = match.group(0)
            is_white_text = bool(re.search(r'color:\s*(?:#ffffff|#fff|rgb\(\s*255\s*,\s*255\s*,\s*255\s*\))', p_tag, re.IGNORECASE))
            if is_white_text and "black-banner-span" not in p_tag and "background-color:#404040" not in p_tag and "table-header-cell" not in p_tag:
                def fix_span(sm):
                    span = sm.group(0)
                    if 'background-color' not in span:
                        span = span.replace('style="', f'style="background-color:{primary_color};')
                    else:
                        span = re.sub(r'background-color:\s*[^;"]+', f'background-color:{primary_color}', span)
                    if 'display:' not in span:
                        span = span.replace('style="', 'style="display:inline-block;padding:2px 6px;border-radius:2px;')
                    return span
                return re.sub(r'<span\s+[^>]*>.*?</span>', fix_span, p_tag, flags=re.DOTALL)
            return p_tag

        cleaned = re.sub(r'<p\s+[^>]*>.*?</p>', format_labelbar_p, cleaned, flags=re.DOTALL)

        # 5. Extract exact images in body region ONLY
        exact_image_html_divs = []
        try:
            img_infos = page.get_image_info(xrefs=True)
            for info in img_infos:
                xref = info.get("xref")
                bbox = info.get("bbox")
                if not xref or not bbox:
                    continue
                x0, y0, x1, y1 = bbox
                if y0 < hy_cutoff or y1 > fy_cutoff:
                    continue
                w_pt = max(1.0, x1 - x0)
                h_pt = max(1.0, y1 - y0)
                try:
                    base_img = doc.extract_image(xref)
                    if not base_img:
                        continue
                    img_bytes = base_img.get("image")
                    img_ext = base_img.get("ext", "png").lower()
                    b64_str = base64.b64encode(img_bytes).decode("utf-8")
                    img_tag = (
                        f'<img src="data:image/{img_ext};base64,{b64_str}" '
                        f'style="position:absolute !important; left:{x0:.1f}pt !important; '
                        f'top:{y0:.1f}pt !important; width:{w_pt:.1f}pt !important; '
                        f'height:{h_pt:.1f}pt !important; z-index:5 !important; '
                        f'object-fit:contain !important;" />'
                    )
                    exact_image_html_divs.append(img_tag)
                except Exception:
                    pass
        except Exception:
            pass

        section_overlays = get_page_section_overlays(page, page_left_str, page_width_str, hy_cutoff, fy_cutoff)

        html_parts.append(cleaned)
        html_parts.extend(vector_html_divs)
        html_parts.extend(section_overlays)
        html_parts.extend(table_header_html_divs)
        html_parts.extend(table_grid_html_divs)
        html_parts.extend(exact_image_html_divs)
        html_parts.append("</div>")

    html_parts.append("</div></body></html>")
    return "\n".join(html_parts)


def generate_dynamic_template_html(data: dict, doc_title: str = "Uploaded Document", theme_config: dict = None) -> str:
    """
    Renders extracted PDF JSON data dynamically into a clean HTML report template.
    Does NOT recreate vendor logos, patient header cards, or footer signatures.
    """
    cfg = get_merged_theme_config(theme_config)
    colors_cfg = cfg.get("colors", {})
    typo_cfg = cfg.get("typography", {})

    primary_color = colors_cfg.get("primary", "#1f497d")
    secondary_color = colors_cfg.get("secondary", "#008080")
    accent_orange = colors_cfg.get("accent_orange", "#ed7d31")
    accent_red = colors_cfg.get("accent_red", "#ff0000")
    bg_page = colors_cfg.get("background_page", "#ffffff")
    text_color = colors_cfg.get("text_primary", "#0d0d0d")
    border_color = colors_cfg.get("border_table", "#000000")
    table_header_bg = primary_color
    table_header_text = colors_cfg.get("text_light", "#ffffff")
    font_family = typo_cfg.get("primary_family", "Cambria, 'Times New Roman', serif")
    result_positive_color = colors_cfg.get("result_positive", "#C00000")
    result_negative_color = colors_cfg.get("result_negative", "#008000")

    spacing_cfg = cfg.get("spacing", {})
    box_padding = spacing_cfg.get("boxPadding", 8)
    cell_padding_x = spacing_cfg.get("cellPaddingX", 6)
    cell_padding_y = spacing_cfg.get("cellPaddingY", 4)
    block_gap = spacing_cfg.get("blockGap", 10)

    border_cfg = cfg.get("border", {})
    border_width = border_cfg.get("width", 1)
    border_style = border_cfg.get("style", "single")
    border_color_val = border_cfg.get("color", "#000000")

    css_border_style = "solid"
    if border_style == "dotted":
        css_border_style = "dotted"
    elif border_style == "dashed":
        css_border_style = "dashed"

    show_tables = theme_config.get("show_tables", True) if theme_config else True
    show_sections = theme_config.get("show_sections", True) if theme_config else True
    show_images = theme_config.get("show_images", True) if theme_config else True
    show_badges = theme_config.get("show_badges", True) if theme_config else True

    badge_rules = (theme_config or {}).get("badge_rules", {
        "danger": ["pathogenic", "positive", "high", "failed", "rejected", "overdue",
                   "invalid", "expired", "critical", "denied", "delinquent", "abnormal"],
        "warning": ["vus", "uncertain", "warning", "pending", "under review",
                    "partial", "provisional", "conditional"],
        "success": ["passed", "normal", "negative", "approved", "paid", "valid",
                    "cleared", "completed", "compliant", "settled"]
    })

    # Keywords that trigger result-level coloring (inline, not badge)
    _POSITIVE_KEYWORDS = ["positive", "pathogenic", "detected", "high", "abnormal", "msi - high", "msi-high"]
    _NEGATIVE_KEYWORDS = ["negative", "normal", "not detected", "stable", "msi - stable", "msi-stable", "benign", "likely benign"]

    def _result_color_span(cell_str: str) -> str:
        """Wrap cell text in a colored span if it matches a clinical result keyword."""
        cell_lower = cell_str.strip().lower()
        if any(kw in cell_lower for kw in _NEGATIVE_KEYWORDS):
            return f"<span style='color:{result_negative_color}; font-weight:bold;'>{cell_str}</span>"
        if any(kw in cell_lower for kw in _POSITIVE_KEYWORDS):
            return f"<span style='color:{result_positive_color}; font-weight:bold;'>{cell_str}</span>"
        return cell_str

    tables = data.get("all_tables") or data.get("tables") or []
    sections = data.get("all_boxes_and_sections") or data.get("content_sections") or []
    images = data.get("all_images_and_graphs") or data.get("images_and_graphs") or []

    # Titles to skip (demographics, generic "General Content / Notes")
    _SKIP_TITLES = {
        "general content / notes",
        "patient details & metadata",
    }

    def _should_skip_section(sec):
        title = sec.get("title", "").strip()
        sec_type = sec.get("type", "")
        if sec_type == "demographics_box":
            return True
        if title.startswith("Header & Metadata Box"):
            return True
        if title.lower() in _SKIP_TITLES:
            return True
        return False

    def _is_reference_fragment(sec):
        """Detect boxes that look like split reference entries (by checking title and content)."""
        title = sec.get("title", "").strip()
        content_text = sec.get("content_text", [])
        body_str = " ".join(content_text) if isinstance(content_text, list) else str(content_text)
        
        combined = (title + " " + body_str).lower()
        
        # Explicit reference cues
        if "pmid" in combined or "doi:" in combined or "et al" in combined:
            return True
        if "http" in combined or "www." in combined or ".html" in combined or ".com/" in combined or "release" in combined:
            return True
        if "guideline" in combined:
            return True
        # Citation year or similar patterns
        if re.search(r'\b\d{4}\b', title):
            return True
        # Starts with a number (like volume info "2;23(8)") and contains punctuation
        if re.match(r'^\d', title) and any(c in title for c in [';', ':', '(', ')', '-']):
            return True
            
        return False

    # 1. Content Section Boxes HTML — merge references, skip generic boxes
    sections_html = ""
    if show_sections and sections:
        reference_paragraphs = []  # collect all reference fragments
        has_references_section = False

        for sec in sections:
            if _should_skip_section(sec):
                continue

            title = sec.get("title", "").strip()
            content_text = sec.get("content_text", [])
            if isinstance(content_text, list):
                body_lines = [t.strip() for t in content_text if t and t.strip()]
            else:
                body_lines = [str(content_text).strip()] if content_text else []

            # Check if this is a "References:" section or a reference fragment
            if title.lower().startswith("references"):
                has_references_section = True
                reference_paragraphs.extend(body_lines)
                continue
            if _is_reference_fragment(sec):
                # Merge title + body into a single reference line
                merged_line = title
                if body_lines:
                    merged_line = f"{title} {' '.join(body_lines)}"
                reference_paragraphs.append(merged_line)
                continue

            if not title and not body_lines:
                continue

            body_paragraphs = "".join([f'<p class="section-p">{t}</p>' for t in body_lines])

            sec_title_html = f'<div class="section-title">{title}</div>' if title else ''
            sections_html += f"""
            <div class="section-box">
                {sec_title_html}
                <div class="section-body">
                    {body_paragraphs}
                </div>
            </div>
            """

        # Emit merged references section at the end
        if reference_paragraphs:
            ref_body = "".join([f'<p class="section-p">{r}</p>' for r in reference_paragraphs])
            sections_html += f"""
            <div class="section-box">
                <div class="section-title">References</div>
                <div class="section-body">
                    {ref_body}
                </div>
            </div>
            """

    # 2. Data Tables HTML — auto-fit by removing empty trailing columns, add result coloring
    tables_html = ""
    if show_tables and tables:
        for t_idx, tab in enumerate(tables):
            headers = list(tab.get("headers", []))
            rows = [list(r) for r in (tab.get("rows", []))]
            page_n = tab.get("page", 1)
            if not headers and not rows:
                continue

            # Determine effective column count: trim empty trailing columns
            raw_ncols = max([len(headers)] + [len(r) for r in rows] + [0])
            if raw_ncols == 0:
                continue
            # Pad to uniform width
            headers = (headers + [""] * raw_ncols)[:raw_ncols]
            rows = [(r + [""] * raw_ncols)[:raw_ncols] for r in rows]

            # Find columns where header AND all row cells are empty → remove them
            keep_cols = []
            for ci in range(raw_ncols):
                col_vals = [headers[ci].strip()] + [str(r[ci]).strip() for r in rows]
                if any(v for v in col_vals):  # at least one non-empty value
                    keep_cols.append(ci)
            if not keep_cols:
                continue

            headers = [headers[ci] for ci in keep_cols]
            rows = [[r[ci] for ci in keep_cols] for r in rows]

            th_html = "".join([f"<th>{h}</th>" for h in headers]) if any(h.strip() for h in headers) else ""
            tr_html = ""
            for r in rows:
                tds = ""
                for cell in r:
                    cell_str = str(cell).replace('\n', '<br>')
                    cell_lower = cell_str.lower()
                    cell_formatted = _result_color_span(cell_str)
                    # Badge logic only if result coloring didn't already wrap it
                    if cell_formatted == cell_str and show_badges:
                        if any(w in cell_lower for w in badge_rules.get("danger", [])):
                            cell_formatted = f"<span class='badge-danger'>{cell_str}</span>"
                        elif any(w in cell_lower for w in badge_rules.get("warning", [])):
                            cell_formatted = f"<span class='badge-warning'>{cell_str}</span>"
                        elif any(w in cell_lower for w in badge_rules.get("success", [])):
                            cell_formatted = f"<span class='badge-success'>{cell_str}</span>"
                    tds += f"<td>{cell_formatted}</td>"
                tr_html += f"<tr>{tds}</tr>"

            table_head_block = f"<thead><tr>{th_html}</tr></thead>" if th_html else ""
            tables_html += f"""
            <div class="table-card-box">
                <div class="table-card-header">
                    Table {t_idx + 1} (Page {page_n})
                </div>
                <table class="table-custom">
                    {table_head_block}
                    <tbody>{tr_html}</tbody>
                </table>
            </div>
            """


    # 5. Images & Graphs Section HTML
    images_html = ""
    if show_images and images:
        img_cards = ""
        for img in images:
            data_uri = img.get("data_uri")
            page_n = img.get("page", 1)
            img_type = img.get("type", "image").replace("_", " ").title()
            w = img.get("width", 0)
            h = img.get("height", 0)
            if not data_uri:
                continue
            img_cards += f"""
            <div class="image-card">
                <img src="{data_uri}" alt="{img_type}" class="extracted-img" />
                <div class="image-caption">Page {page_n} • {img_type} ({w}×{h}px)</div>
            </div>
            """
        if img_cards:
            images_html = f"""
            <div class="section-box" style="margin-top: 14pt;">
                <div class="section-title">Extracted Body Images & Graphical Content</div>
                <div class="image-grid">
                    {img_cards}
                </div>
            </div>
            """

    oncquest_logo_html = ""
    patient_table_html = ""

    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>{doc_title}</title>
<style>
@page {{ size: 595.6pt 842.0pt; margin: 0; }}
* {{ box-sizing: border-box; }}
body {{ margin: 0; padding: 0; background-color: #f1f5f9; font-family: {font_family}; color: {text_color}; }}
.pdf-container {{ display: flex; flex-direction: column; align-items: center; padding: 20px 0; }}
.report-content {{ background: {bg_page}; width: 595.6pt; padding: 35.5pt; margin-bottom: 20px; position: relative; box-shadow: 0 4px 12px rgba(0,0,0,0.15); font-family: {font_family}; word-break: normal; overflow-wrap: break-word; }}
.section-box {{ border: {border_width}px {css_border_style} {border_color_val}; margin-bottom: {block_gap}pt; padding: {box_padding}px; width: fit-content; max-width: 100%; box-sizing: border-box; position: relative; background: #ffffff; overflow: visible; }}
.section-title {{ color: {primary_color}; font-weight: bold; font-size: 10.5pt; padding: 6pt 10pt; display: block; letter-spacing: 0.02em; border-bottom: 2px solid {primary_color}; }}
.section-body {{ padding: 10pt 12pt; font-size: 9.5pt; line-height: 1.5; color: {text_color}; }}
.section-p {{ margin: 0 0 6pt 0; text-align: left; word-break: normal; overflow-wrap: break-word; }}
.section-p:last-child {{ margin-bottom: 0; }}
.table-card-box {{ margin-top: 14pt; margin-bottom: {block_gap}pt; border: {border_width}px {css_border_style} {border_color_val}; width: fit-content; max-width: 100%; box-sizing: border-box; overflow: visible; background: #ffffff; }}
.table-card-header {{ font-size: 9.5pt; font-weight: bold; color: {primary_color}; padding: 6pt 10pt; background-color: #f8fafc; border-bottom: 1px solid {border_color_val}; }}
.table-custom {{ border-collapse: collapse; font-size: 9.5pt; table-layout: auto; width: 100%; }}
.table-custom th {{ background-color: {table_header_bg}; color: {table_header_text}; font-family: {font_family}; font-size: 9.5pt; font-weight: bold; text-align: center; padding: {cell_padding_y}px {cell_padding_x}px; border: {border_width}px {css_border_style} {border_color_val}; word-break: normal; overflow-wrap: break-word; }}
.table-custom td {{ padding: {cell_padding_y}px {cell_padding_x}px; border: {border_width}px {css_border_style} {border_color_val}; vertical-align: top; line-height: 1.35; font-size: 9.5pt; word-break: normal; overflow-wrap: break-word; }}
.badge-danger {{ background: #dc2626; color: #ffffff; padding: 2px 6px; border-radius: 3px; font-weight: bold; display: inline-block; font-size: 8.5pt; }}
.badge-warning {{ background: #d97706; color: #ffffff; padding: 2px 6px; border-radius: 3px; font-weight: bold; display: inline-block; font-size: 8.5pt; }}
.badge-success {{ background: #16a34a; color: #ffffff; padding: 2px 6px; border-radius: 3px; font-weight: bold; display: inline-block; font-size: 8.5pt; }}
.image-grid {{ display: flex; flex-wrap: wrap; gap: 12px; margin-top: 8pt; padding: 10pt; }}
.image-card {{ border: 1px solid #e2e8f0; border-radius: 4px; padding: 8px; background: #fafafa; text-align: center; max-width: 240px; }}
.extracted-img {{ max-width: 100%; max-height: 160px; object-fit: contain; border-radius: 2px; }}
.image-caption {{ font-size: 8pt; color: #64748b; margin-top: 4pt; font-family: sans-serif; }}
</style>
</head>
<body>
<div class="pdf-container">
  <div class="report-content">
    {sections_html}

    {tables_html}

    {images_html}
  </div>
</div>
</body>
</html>"""
    return html_content




def _set_cell_background(cell, hex_color: str):
    """Utility to set XML background shading of a python-docx table cell."""
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    hex_clean = hex_color.lstrip("#")
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_clean}"/>')
    cell._tc.get_or_add_tcPr().append(shd)


def _set_cell_borders(cell, border_color: str = "CBD5E1", border_size: str = "4"):
    """Utility to set XML cell borders of a python-docx table cell."""
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    hex_clean = border_color.lstrip("#")
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="{border_size}" w:space="0" w:color="{hex_clean}"/>'
        f'<w:left w:val="single" w:sz="{border_size}" w:space="0" w:color="{hex_clean}"/>'
        f'<w:bottom w:val="single" w:sz="{border_size}" w:space="0" w:color="{hex_clean}"/>'
        f'<w:right w:val="single" w:sz="{border_size}" w:space="0" w:color="{hex_clean}"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)


def _set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    """Utility to set XML cell margins/padding of a python-docx table cell in dxa (1pt = 20dxa)."""
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)


def _set_cell_width(cell, width_pt):
    """Utility to set XML cell width in dxa (1pt = 20dxa)."""
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    tcPr = cell._tc.get_or_add_tcPr()
    width_dxa = int(width_pt * 20)
    tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{width_dxa}" w:type="dxa"/>')
    tcPr.append(tcW)


def _add_page_number_to_run(run, docx_module):
    """Utility to add dynamic Word PAGE field to a run."""
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    fldChar1 = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>')
    instrText = parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> PAGE </w:instrText>')
    fldChar2 = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="separate"/>')
    fldChar3 = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>')
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)


def _add_numpages_to_run(run, docx_module):
    """Utility to add dynamic Word NUMPAGES field to a run."""
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    fldChar1 = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>')
    instrText = parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> NUMPAGES </w:instrText>')
    fldChar2 = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="separate"/>')
    fldChar3 = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>')
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)


def _hex_to_rgb(hex_str: str):
    from docx.shared import RGBColor
    if not hex_str:
        return RGBColor(31, 73, 125)
    hex_clean = str(hex_str).lstrip("#")
    if len(hex_clean) == 6:
        r, g, b = int(hex_clean[0:2], 16), int(hex_clean[2:4], 16), int(hex_clean[4:6], 16)
    else:
        r, g, b = 31, 73, 125
    return RGBColor(r, g, b)


def get_merged_theme_config(theme_config: dict = None) -> dict:
    """
    Dynamically loads styling and layout parameters from theme.json and merges them with user overrides.
    """
    merged = {
        "document_page": {
            "paper_size": "A4",
            "width_pt": 595.6,
            "height_pt": 842.0,
            "margins_pt": {"top": 36.0, "bottom": 36.0, "left": 36.0, "right": 36.0},
            "header_distance_pt": 18.0,
            "footer_distance_pt": 18.0
        },
        "typography": {
            "primary_family": "Cambria",
            "secondary_family": "Calibri",
            "line_spacing": 1.15,
            "paragraph_space_after_pt": 4.0,
            "sizes_pt": {
                "document_title": 14.0,
                "section_heading": 13.0,
                "banner_heading": 10.0,
                "body": 10.0,
                "table_header": 9.5,
                "table_body": 9.5,
                "footer": 9.0,
                "small": 8.0
            }
        },
        "colors": {
            "primary": "#1f497d",
            "secondary": "#008080",
            "accent_orange": "#ed7d31",
            "accent_red": "#ff0000",
            "banner_dark": "#404040",
            "text_primary": "#000000",
            "text_secondary": "#242729",
            "text_light": "#ffffff",
            "background_page": "#ffffff",
            "border_primary": "#1f497d",
            "border_table": "#000000",
            "border_header": "#cbd5e1",
            "alternating_row_bg": "#f8fafc",
            "result_positive": "#C00000",
            "result_negative": "#008000"
        },
        "headers_and_footers": {
            "header": {
                "show_logo": True,
                "logo_image_path": "assets/header_image1.png",
                "logo_width_pt": 540.0,
                "logo_alignment": "center",
                "show_metadata_table": True,
                "metadata_table": {
                    "border_color": "#cbd5e1",
                    "border_size": "2",
                    "label_font_weight": "bold",
                    "value_font_weight": "normal",
                    "column_widths_pt": [87.2, 160.75, 88.95, 203.1],
                    "cell_padding_pt": {"top": 2, "bottom": 2, "left": 4, "right": 4}
                }
            },
            "footer": {
                "show_signatures": True,
                "signature_image_path": "assets/footer_signatures.png",
                "signature_width_pt": 540.0,
                "signature_alignment": "center",
                "show_page_number": True,
                "page_number_text": "Page {page} of {total}",
                "font_size_pt": 9.0,
                "text_color": "#404040"
            }
        },
        "spacing": {
            "boxPadding": 8,
            "cellPaddingX": 6,
            "cellPaddingY": 4,
            "blockGap": 10
        },
        "border": {
            "width": 1,
            "style": "single",
            "color": "#000000"
        },
        "components": {
            "black_banner": {
                "background_color": "#404040",
                "text_color": "#ffffff",
                "font_family": "Cambria",
                "font_size_pt": 13.0,
                "font_weight": "bold",
                "alignment": "center",
                "space_before_pt": 6.0,
                "space_after_pt": 6.0
            },
            "blue_section_banner": {
                "background_color": "#1f497d",
                "text_color": "#ffffff",
                "font_family": "Cambria",
                "font_size_pt": 10.0,
                "font_weight": "bold",
                "alignment": "left",
                "space_before_pt": 4.0,
                "space_after_pt": 4.0
            },
            "data_table": {
                "header_background_color": "#1f497d",
                "header_text_color": "#ffffff",
                "header_font_weight": "bold",
                "header_font_size_pt": 9.5,
                "body_font_size_pt": 9.5,
                "border_color": "#1f497d",
                "border_size": "4",
                "cell_padding_pt": {"top": 3, "bottom": 3, "left": 4, "right": 4},
                "alignment": "center"
            },
            "content_box": {
                "border_color": "#1f497d",
                "border_size": "4",
                "background_color": "transparent",
                "title_font_size_pt": 10.0,
                "title_font_weight": "bold",
                "title_text_color": "#1f497d",
                "body_font_size_pt": 9.5,
                "body_text_color": "#000000",
                "space_before_pt": 4.0,
                "space_after_pt": 4.0
            },
            "bullet_list": {
                "bullet_char": "•",
                "indent_pt": 18.0,
                "font_size_pt": 9.5
            },
            "end_of_report_marker": {
                "text": "------------------ End Of Report ------------------",
                "font_family": "Calibri",
                "font_size_pt": 10.0,
                "text_color": "#000000",
                "alignment": "center",
                "space_before_pt": 12.0,
                "space_after_pt": 12.0
            }
        },
        "word": {
            "layout": {
                "element_flow": ["header_metadata", "headings", "banners", "content_boxes", "tables", "paragraphs", "bullet_lists", "signatures", "end_of_report_marker"],
                "spacing": {
                    "default_line_spacing": 1.15,
                    "paragraph_space_before_pt": 0.0,
                    "paragraph_space_after_pt": 4.0,
                    "heading_space_before_pt": 8.0,
                    "heading_space_after_pt": 4.0,
                    "table_space_before_pt": 6.0,
                    "table_space_after_pt": 6.0,
                    "content_box_space_before_pt": 6.0,
                    "content_box_space_after_pt": 6.0,
                    "banner_space_before_pt": 6.0,
                    "banner_space_after_pt": 6.0
                },
                "boxes": {
                    "border_color": "#1f497d",
                    "border_size": "4",
                    "background_color": "#ffffff",
                    "title_font_size_pt": 10.0,
                    "title_font_weight": "bold",
                    "title_text_color": "#1f497d",
                    "body_font_size_pt": 9.5,
                    "body_text_color": "#000000",
                    "cell_padding_pt": {"top": 4, "bottom": 4, "left": 6, "right": 6}
                },
                "tables": {
                    "alignment": "center",
                    "header_background_color": "#1f497d",
                    "header_text_color": "#ffffff",
                    "header_font_weight": "bold",
                    "header_font_size_pt": 9.5,
                    "body_font_size_pt": 9.5,
                    "border_color": "#1f497d",
                    "border_size": "4",
                    "alternating_row_bg": "#f8fafc",
                    "cell_padding_pt": {"top": 3, "bottom": 3, "left": 4, "right": 4},
                    "autofit": True,
                    "cant_split_rows": True
                },
                "banners": {
                    "black_banner": {"background_color": "#404040", "text_color": "#ffffff", "font_size_pt": 13.0, "font_weight": "bold", "alignment": "center"},
                    "blue_banner": {"background_color": "#1f497d", "text_color": "#ffffff", "font_size_pt": 10.0, "font_weight": "bold", "alignment": "left"}
                }
            },
            "pagination": {
                "respect_page_breaks": True,
                "keep_with_next_headings": True,
                "page_numbering": {"enabled": True, "text_format": "Page {page} of {total}", "font_size_pt": 9.0, "alignment": "center"},
                "repeat_table_headers": True,
                "prevent_orphan_rows": True
            }
        },
        "primary_color": "#1f497d",
        "secondary_color": "#008080",
        "accent_orange": "#ed7d31",
        "accent_red": "#ff0000",
        "banner_dark": "#404040",
        "font_family": "Cambria",
        "border_color": "#1f497d",
        "body_font_size": 10.0,
        "title_font_size": 14.0,
        "table_header_font_size": 9.5,
        "show_footer_signatures": True,
        "end_report_marker": "------------------ End Of Report ------------------"
    }

    theme_file = Path("theme.json")
    if theme_file.exists():
        try:
            with open(theme_file, "r", encoding="utf-8") as f:
                tj = json.load(f)
                if isinstance(tj, dict):
                    for key in ["document_page", "typography", "colors", "headers_and_footers", "components", "word", "spacing", "border"]:
                        if key in tj and isinstance(tj[key], dict):
                            if key in merged and isinstance(merged[key], dict):
                                merged[key].update(tj[key])
                            else:
                                merged[key] = tj[key]

                    colors = merged.get("colors", {})
                    fonts = merged.get("typography", {})
                    if "primary" in colors:
                        merged["primary_color"] = colors["primary"]
                    if "secondary" in colors:
                        merged["secondary_color"] = colors["secondary"]
                    if "accent_orange" in colors:
                        merged["accent_orange"] = colors["accent_orange"]
                    if "accent_red" in colors:
                        merged["accent_red"] = colors["accent_red"]
                    if "banner_dark" in colors:
                        merged["banner_dark"] = colors["banner_dark"]
                    if "border_primary" in colors:
                        merged["border_color"] = colors["border_primary"]

                    merged["font_family"] = fonts.get("primary_family", "Cambria")
                    sizes = fonts.get("sizes_pt", {})
                    if "body" in sizes:
                        merged["body_font_size"] = sizes["body"]
                    if "document_title" in sizes:
                        merged["title_font_size"] = sizes["document_title"]
                    if "table_header" in sizes:
                        merged["table_header_font_size"] = sizes["table_header"]

                    end_marker_comp = merged.get("components", {}).get("end_of_report_marker", {})
                    if "text" in end_marker_comp:
                        merged["end_report_marker"] = end_marker_comp["text"]
        except Exception:
            pass

    if theme_config and isinstance(theme_config, dict):
        for k, v in theme_config.items():
            if v is not None:
                merged[k] = v

    return merged


def _find_kv_val(kv_dict: dict, keywords: list) -> str:
    """Helper to fuzzy match key in key-value dictionary."""
    if not isinstance(kv_dict, dict):
        return ""
    for k, v in kv_dict.items():
        k_clean = str(k).lower().replace("_", " ").replace(".", " ").strip()
        for kw in keywords:
            if kw in k_clean:
                return str(v).strip()
    return ""


def clean_extracted_text(val):
    """
    Clean trailing/standalone '_' artifacts and stray whitespace from extracted text and table cells.
    Preserves internal underscores (e.g. 'Gene_Name', 'NM_000251.3').
    """
    if not isinstance(val, str):
        return val
    text = val.strip()
    if not text:
        return ""
    # If text consists solely of underscores and whitespace
    if set(text) <= {'_', ' '}:
        return ""
    # Remove trailing/standalone underscore artifacts like " _", " _ _", or trailing "_" preceded by whitespace
    text = re.sub(r'(?:\s+_+)+\s*$', '', text)
    return text.strip()


def is_decorative_or_watermark_image(img_dict: dict, page_w: float = 595.6, page_h: float = 842.0) -> bool:
    """
    Filter out decorative/background images: skip images that are near-square and very large relative to page,
    or that sit in page margins/watermark layer, or have low content density.
    Keep only real content graphs/logos based on size & aspect-ratio thresholds.
    """
    if not isinstance(img_dict, dict):
        return False

    bbox = img_dict.get("bbox") or [0, 0, 0, 0]
    x0, y0, x1, y1 = bbox[:4] if len(bbox) >= 4 else (0, 0, 0, 0)
    w_pt = x1 - x0 if x1 > x0 else 0.0
    h_pt = y1 - y0 if y1 > y0 else 0.0

    px_w = float(img_dict.get("width") or 0.0)
    px_h = float(img_dict.get("height") or 0.0)

    data_uri = img_dict.get("data_uri")
    if data_uri and (px_w == 0 or px_h == 0 or w_pt == 0):
        try:
            b64_str = data_uri.split(",", 1)[1] if "," in data_uri else data_uri
            img_bytes = base64.b64decode(b64_str)
            with Image.open(io.BytesIO(img_bytes)) as pil_img:
                px_w, px_h = float(pil_img.size[0]), float(pil_img.size[1])
        except Exception:
            pass

    eff_w = w_pt if w_pt > 0 else px_w
    eff_h = h_pt if h_pt > 0 else px_h

    if eff_w <= 0 or eff_h <= 0:
        return False

    aspect_ratio = eff_w / eff_h if eff_h > 0 else 1.0

    # Criteria 1: Near-square (0.7 <= aspect <= 1.4) AND large relative to page
    is_near_square = 0.7 <= aspect_ratio <= 1.4
    is_large_pt = (w_pt > 350 and h_pt > 350) or (w_pt > 0.55 * page_w and h_pt > 0.4 * page_h)
    is_large_px = (px_w > 450 and px_h > 450)

    if is_near_square and (is_large_pt or is_large_px):
        return True

    # Criteria 2: Full-page background or margin-spanning watermark layer
    if w_pt > 0.8 * page_w and h_pt > 0.8 * page_h:
        return True
    if w_pt > 0 and h_pt > 0 and x0 <= 40 and x1 >= (page_w - 40) and y0 <= 50 and y1 >= (page_h - 50):
        return True

    return False


def _clean_text(s):
    if s is None:
        return ""
    s = str(s)
    s = re.sub(r"\s*_+\s*", " ", s)   # remove standalone/trailing underscores
    s = re.sub(r"\s{2,}", " ", s)     # collapse whitespace
    return s.strip()


def _is_decorative_image(img, page_w=595.0, page_h=842.0):
    w = float(img.get("width", 0) or 0)
    h = float(img.get("height", 0) or 0)
    if w <= 0 or h <= 0:
        return False
    aspect = w / h if h else 1.0
    near_square = 0.7 <= aspect <= 1.4
    bb = img.get("bbox")
    if isinstance(bb, (list, tuple)) and len(bb) >= 4:
        bw = abs(bb[2] - bb[0]); bh = abs(bb[3] - bb[1])
        area_frac = (bw * bh) / (page_w * page_h)
    else:
        area_frac = (w * h) / (page_w * page_h)
    return near_square and area_frac >= 0.22


def _load_oncquest_theme(theme_config=None):
    theme_config = theme_config or {}
    tj = {}
    tj_path = Path("theme.json")
    if tj_path.exists():
        try:
            with open(tj_path, "r", encoding="utf-8") as f:
                tj = json.load(f)
        except Exception:
            tj = {}
 
    colors = tj.get("colors", {}) or {}
    text = colors.get("text", {}) or {}
    fonts = tj.get("fonts", {}) or {}
    families = fonts.get("families", {}) or {}
    sizes = fonts.get("sizes", {}) or {}
    comps = tj.get("components", {}) or {}
    eor = comps.get("end_of_report_marker", {}) or {}

    def clean_font(css, fallback="Cambria"):
        if not css or not isinstance(css, str):
            return fallback
        first = css.split(",")[0].strip().strip("'\"")
        if first.lower() in ("bold", "italic", "bolditalic", "serif", "sans-serif"):
            return fallback
        return first or fallback

    def clean_hex(h, fallback):
        if not h or not isinstance(h, str):
            return fallback
        h = h.strip().lstrip("#")
        return h if len(h) == 6 else fallback

    border = colors.get("border")
    border_primary = border.get("primary") if isinstance(border, dict) else None
    primary = theme_config.get("primary_color") or colors.get("primary") or "#1f497d"

    return {
        "primary": clean_hex(primary, "1f497d"),
        "banner_dark": clean_hex(colors.get("banner_dark"), "404040"),
        "border": clean_hex(border_primary, "1f497d"),
        "border_table": clean_hex(colors.get("border_table"), "000000"),
        "body_color": clean_hex(text.get("primary"), "000000"),
        "header_text": clean_hex(text.get("white"), "ffffff"),
        "result_positive": clean_hex(colors.get("result_positive"), "C00000"),
        "result_negative": clean_hex(colors.get("result_negative"), "008000"),
        "font": clean_font(families.get("primary"), "Cambria"),
        "table_header_font": clean_font(families.get("table_header"), "Cambria"),
        "body_pt": float(sizes.get("body_pt", 10.0)),
        "banner_pt": float(sizes.get("banner_pt", 13.0)),
        "table_header_pt": float(sizes.get("table_header_pt", 9.5)),
        "eor_text": eor.get("text_pattern",
                            "------------------ End Of Report ------------------"),
        "eor_font": clean_font(eor.get("font_family"), "Calibri"),
        "eor_pt": float(eor.get("font_size_pt", 10.0)),
        "eor_color": clean_hex(eor.get("text_color"), "000000"),
    }


def convert_json_to_docx(data: dict, output_path: str = None, theme_config: dict = None):
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    T = _load_oncquest_theme(theme_config)

    def rgb(hx):
        return RGBColor(int(hx[0:2], 16), int(hx[2:4], 16), int(hx[4:6], 16))

    # ---------- xml helpers ----------
    def shade_cell(cell, fill):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), fill); tcPr.append(shd)

    def shade_para(paragraph, fill):
        pPr = paragraph._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), fill); pPr.append(shd)

    def cell_margins(cell, t=40, b=40, l=80, r=80):
        tcPr = cell._tc.get_or_add_tcPr()
        m = OxmlElement("w:tcMar")
        for tag, val in (("top", t), ("bottom", b), ("start", l), ("end", r)):
            n = OxmlElement(f"w:{tag}")
            n.set(qn("w:w"), str(val)); n.set(qn("w:type"), "dxa"); m.append(n)
        tcPr.append(m)

    def table_borders(table, color, sz=4):
        tblPr = table._tbl.tblPr
        b = OxmlElement("w:tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            e = OxmlElement(f"w:{edge}")
            e.set(qn("w:val"), "single"); e.set(qn("w:sz"), str(sz))
            e.set(qn("w:space"), "0"); e.set(qn("w:color"), color); b.append(e)
        tblPr.append(b)

    def set_repeat_header(row):
        """Make a table's header row repeat on every page."""
        trPr = row._tr.get_or_add_trPr()
        th = OxmlElement("w:tblHeader"); th.set(qn("w:val"), "true")
        trPr.append(th)

    def set_cant_split(row):
        """Prevent a row from splitting across pages."""
        trPr = row._tr.get_or_add_trPr()
        cs = OxmlElement("w:cantSplit"); cs.set(qn("w:val"), "true")
        trPr.append(cs)

    # Clinical result keywords for coloring
    _POS_KW = ["positive", "pathogenic", "detected", "high", "abnormal", "msi - high", "msi-high"]
    _NEG_KW = ["negative", "normal", "not detected", "stable", "msi - stable", "msi-stable", "benign", "likely benign"]

    def _result_color(cell_text):
        """Return RGBColor for clinical result or None."""
        cl = cell_text.strip().lower()
        if any(kw in cl for kw in _NEG_KW):
            return rgb(T["result_negative"])
        if any(kw in cl for kw in _POS_KW):
            return rgb(T["result_positive"])
        return None

    # ---------- render primitives ----------
    def add_banner(doc, txt, fill, big=False, flat=False):
        """Add a heading banner. flat=True renders colored text without background shading."""
        txt = _clean_text(txt)
        if not txt:
            return
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True  # heading stays with its content
        if not flat:
            shade_para(p, fill)
        run = p.add_run(txt); run.bold = True
        run.font.name = T["font"]
        run.font.size = Pt(T["banner_pt"] if big else T["body_pt"] + 1.0)
        if flat:
            run.font.color.rgb = rgb(T["primary"])
        else:
            run.font.color.rgb = rgb(T["header_text"])

    def add_para(doc, txt):
        txt = _clean_text(txt)
        if not txt:
            return
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.15
        run = p.add_run(txt)
        run.font.name = T["font"]; run.font.size = Pt(T["body_pt"])
        run.font.color.rgb = rgb(T["body_color"])

    def add_kv_table(doc, kv):
        if not kv:
            return
        tbl = doc.add_table(rows=0, cols=2)
        table_borders(tbl, T["border_table"])
        tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
        tbl.autofit = True
        for k, v in kv.items():
            cells = tbl.add_row().cells
            set_cant_split(tbl.rows[-1])
            cell_margins(cells[0]); cell_margins(cells[1])
            kr = cells[0].paragraphs[0].add_run(_clean_text(k))
            kr.bold = True; kr.font.name = T["font"]; kr.font.size = Pt(T["body_pt"])
            kr.font.color.rgb = rgb(T["primary"])
            vr = cells[1].paragraphs[0].add_run(_clean_text(v))
            vr.font.name = T["font"]; vr.font.size = Pt(T["body_pt"])
            vr.font.color.rgb = rgb(T["body_color"])
        doc.add_paragraph().paragraph_format.space_after = Pt(2)

    def add_data_table(doc, tab):
        headers = [_clean_text(h) for h in (tab.get("headers") or [])]
        rows = [[_clean_text(c) for c in (r or [])] for r in (tab.get("rows") or [])]

        if (not headers or not any(headers)) and rows:  # empty header -> promote row 0
            headers = rows[0]; rows = rows[1:]
        if not any(headers) and not any(any(c for c in r) for r in rows):  # empty table
            return

        ncols = max([len(headers)] + [len(r) for r in rows] + [0])
        if ncols == 0:
            return
        headers = (headers + [""] * ncols)[:ncols]
        rows = [(r + [""] * ncols)[:ncols] for r in rows]

        # Remove entirely-empty columns (header + all rows empty)
        keep_cols = []
        for ci in range(ncols):
            col_vals = [headers[ci].strip()] + [str(r[ci]).strip() for r in rows]
            if any(v for v in col_vals):
                keep_cols.append(ci)
        if not keep_cols:
            return
        headers = [headers[ci] for ci in keep_cols]
        rows = [[r[ci] for ci in keep_cols] for r in rows]
        ncols = len(keep_cols)

        has_header = any(headers)
        tbl = doc.add_table(rows=1 if has_header else 0, cols=ncols)
        table_borders(tbl, T["border_table"])
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = True

        if has_header:
            hrow = tbl.rows[0]
            set_repeat_header(hrow)   # header repeats across pages
            set_cant_split(hrow)
            for i in range(ncols):
                cell = hrow.cells[i]
                shade_cell(cell, T["primary"]); cell_margins(cell)
                para = cell.paragraphs[0]; para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = para.add_run(headers[i]); run.bold = True
                run.font.name = T["table_header_font"]
                run.font.size = Pt(T["table_header_pt"])
                run.font.color.rgb = rgb(T["header_text"])

        for r in rows:
            cells = tbl.add_row().cells
            set_cant_split(tbl.rows[-1])   # a row never splits across pages
            for i in range(ncols):
                cell_margins(cells[i])
                cell_text = r[i]
                run = cells[i].paragraphs[0].add_run(cell_text)
                run.font.name = T["font"]; run.font.size = Pt(T["body_pt"])
                # Apply clinical result coloring
                rc = _result_color(cell_text)
                if rc:
                    run.font.color.rgb = rc
                    run.bold = True
                else:
                    run.font.color.rgb = rgb(T["body_color"])
        doc.add_paragraph().paragraph_format.space_after = Pt(2)

    def add_image(doc, img):
        if _is_decorative_image(img):       # skip watermark circles
            return
        uri = img.get("data_uri", "")
        if not uri or "," not in uri:
            return
        try:
            raw = base64.b64decode(uri.split(",", 1)[1])
            stream = io.BytesIO(raw)
            w = float(img.get("width", 0) or 0)
            p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run()
            width_in = min(6.0, w / 96.0) if w else 4.0
            run.add_picture(stream, width=Inches(max(1.0, width_in)))
        except Exception:
            pass

    def bbox_top(el):
        bb = el.get("bbox")
        if isinstance(bb, (list, tuple)) and len(bb) >= 2:
            return (bb[1], bb[0])   # sort by top-Y then left-X
        return (10_000_000, 0)

    def content_lines(box):
        ct = box.get("content_text", [])
        if isinstance(ct, str):
            return [ct]
        if isinstance(ct, list):
            return [str(x) for x in ct]
        return []

    # ---------------- build document ----------------
    doc = Document()

    normal = doc.styles["Normal"]
    normal.font.name = T["font"]; normal.font.size = Pt(T["body_pt"])
    rpr = normal.element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn("w:ascii"), T["font"]); rfonts.set(qn("w:hAnsi"), T["font"])

    summary = data.get("document_summary", {}) or {}
    title = os.path.splitext(summary.get("file_name", "Report"))[0]
    add_banner(doc, title.upper(), T["banner_dark"], big=True)

    kv = data.get("all_key_value_pairs") or data.get("extracted_key_value_pairs") or {}
    if kv:
        add_banner(doc, "DETAILS", T["primary"], flat=True)
        add_kv_table(doc, kv)

    seen = set()

    # Titles to skip in DOCX
    _SKIP_TITLES_DOCX = {
        "general content / notes",
        "patient details & metadata",
    }

    def _should_skip_docx(el):
        ttl = _clean_text(el.get("title"))
        sec_type = el.get("type", "")
        if sec_type == "demographics_box":
            return True
        if ttl and ttl.startswith("Header & Metadata Box"):
            return True
        if ttl and ttl.lower() in _SKIP_TITLES_DOCX:
            return True
        return False

    def _is_ref_fragment_docx(el):
        ttl = _clean_text(el.get("title"))
        content_text = content_lines(el)
        body_str = " ".join(content_text)
        
        combined = (ttl + " " + body_str).lower()
        
        if "pmid" in combined or "doi:" in combined or "et al" in combined:
            return True
        if "http" in combined or "www." in combined or ".html" in combined or ".com/" in combined or "release" in combined:
            return True
        if "guideline" in combined:
            return True
        if re.search(r'\b\d{4}\b', ttl):
            return True
        if re.match(r'^\d', ttl) and any(c in ttl for c in [';', ':', '(', ')', '-']):
            return True
        return False

    ref_lines_docx = []  # collect references for merging

    def render_page(page):
        boxes = page.get("boxes_and_sections") or []
        tables = page.get("tables") or []
        images = page.get("images_and_graphs") or []
        text_blocks = page.get("text_blocks") or []

        items = []
        for b in boxes:
            items.append(("box", bbox_top(b), b))
        for t in tables:
            items.append(("table", bbox_top(t), t))
        for im in images:
            items.append(("image", bbox_top(im), im))
        if not boxes and text_blocks:      # fallback so nothing is lost
            for tb in text_blocks:
                items.append(("textblock", bbox_top(tb), tb))

        items.sort(key=lambda x: x[1])     # bbox used ONLY for ordering

        for kind, _pos, el in items:
            if kind == "box":
                if _should_skip_docx(el):
                    # Still capture content from skipped boxes (they may have useful text)
                    for line in content_lines(el):
                        s = _clean_text(line)
                        if s and s not in seen:
                            seen.add(s); add_para(doc, s)
                    continue
                ttl = _clean_text(el.get("title"))
                # Handle references
                if ttl and ttl.lower().startswith("references"):
                    for line in content_lines(el):
                        s = _clean_text(line)
                        if s:
                            ref_lines_docx.append(s)
                    continue
                if _is_ref_fragment_docx(el):
                    merged = ttl
                    body = [_clean_text(l) for l in content_lines(el) if _clean_text(l)]
                    if body:
                        merged = f"{ttl} {' '.join(body)}"
                    ref_lines_docx.append(merged)
                    continue
                if ttl:
                    add_banner(doc, ttl, T["primary"], flat=True)
                for line in content_lines(el):
                    s = _clean_text(line)
                    if s and s not in seen:
                        seen.add(s); add_para(doc, s)
            elif kind == "table":
                add_data_table(doc, el)
            elif kind == "image":
                add_image(doc, el)
            elif kind == "textblock":
                s = _clean_text(el.get("text", ""))
                if s and s not in seen:
                    seen.add(s); add_para(doc, s)

    pages = data.get("pages") or []
    if pages:
        for page in pages:
            render_page(page)
    else:
        for sec in data.get("all_boxes_and_sections") or data.get("content_sections") or []:
            if _should_skip_docx(sec):
                continue
            ttl = _clean_text(sec.get("title"))
            if ttl and ttl.lower().startswith("references"):
                for line in content_lines(sec):
                    s = _clean_text(line)
                    if s:
                        ref_lines_docx.append(s)
                continue
            if _is_ref_fragment_docx(sec):
                merged = ttl
                body = [_clean_text(l) for l in content_lines(sec) if _clean_text(l)]
                if body:
                    merged = f"{ttl} {' '.join(body)}"
                ref_lines_docx.append(merged)
                continue
            if ttl:
                add_banner(doc, ttl, T["primary"], flat=True)
            for line in content_lines(sec):
                s = _clean_text(line)
                if s and s not in seen:
                    seen.add(s); add_para(doc, s)
        for tb in data.get("all_tables") or data.get("tables") or []:
            add_data_table(doc, tb)
        for im in data.get("all_images_and_graphs") or data.get("images_and_graphs") or []:
            add_image(doc, im)

    # Emit merged references section
    if ref_lines_docx:
        add_banner(doc, "References", T["primary"], flat=True)
        for rl in ref_lines_docx:
            if rl not in seen:
                seen.add(rl); add_para(doc, rl)

    end = doc.add_paragraph(); end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    er = end.add_run(T["eor_text"])
    er.font.name = T["eor_font"]; er.font.size = Pt(T["eor_pt"])
    er.font.color.rgb = rgb(T["eor_color"])

    if output_path:
        out_p = Path(output_path)
        out_p.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(out_p))
        return None
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def render_json_file_to_html(json_path, output_path: str = None, theme_config: dict = None) -> str:
    """
    Renders an extracted document JSON file (containing page text_blocks, drawings, images)
    directly into an HTML document template matching the Oncquest design system.
    Saves to output_path if provided, and returns the full HTML string.
    """
    import html
    json_p = Path(json_path)
    with open(json_p, "r", encoding="utf-8") as f:
        data = json.load(f)

    source_file = data.get("source_file", json_p.stem + ".pdf")
    doc_title = Path(source_file).stem
    pages = data.get("pages", [])

    cfg = get_merged_theme_config(theme_config)
    colors_cfg = cfg.get("colors", {})
    primary_color = colors_cfg.get("primary", "#1f497d")
    secondary_color = colors_cfg.get("secondary", "#008080")
    accent_orange = colors_cfg.get("accent_orange", "#ed7d31")
    accent_red = colors_cfg.get("accent_red", "#ff0000")

    def sanitize_text(text: str) -> str:
        if not text:
            return ""
        text = text.replace("\ufb01", "fi").replace("\ufb02", "fl").replace("\xa0", " ")
        return html.escape(text)

    def get_font_family(font_name: str) -> str:
        if not font_name:
            return "Calibri, sans-serif"
        fn = font_name.lower()
        if "calibri" in fn:
            return "Calibri, Arial, sans-serif"
        elif "tahoma" in fn:
            return "Tahoma, Geneva, sans-serif"
        elif "verdana" in fn:
            return "Verdana, Geneva, sans-serif"
        elif "times" in fn:
            return "'Times New Roman', Times, serif"
        elif "cambria" in fn:
            return "Cambria, Georgia, serif"
        return "Calibri, Arial, sans-serif"

    html_parts = [
        "<!DOCTYPE html>",
        "<html lang='en'>",
        "<head>",
        "<meta charset='utf-8'>",
        "<meta name='viewport' content='width=device-width, initial-scale=1.0'>",
        f"<title>{doc_title} — Oncquest Lab Report</title>",
        "<link href='https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap' rel='stylesheet'>",
        "<style>",
        f"""
        :root {{
          --color-primary: {primary_color};
          --color-secondary: {secondary_color};
          --color-banner-dark: #404040;
          --color-bg-container: #525659;
          --color-bg-page: #ffffff;
          --font-primary: 'Calibri', 'Inter', sans-serif;
          --page-shadow: 0 8px 24px rgba(0, 0, 0, 0.25);
        }}
        * {{ box-sizing: border-box; margin: 0; padding: 0; }}
        body {{
          margin: 0;
          padding: 0;
          background-color: var(--color-bg-container);
          font-family: var(--font-primary);
          color: #000000;
          -webkit-font-smoothing: antialiased;
        }}
        .header-bar {{
          background: #0f172a;
          color: #ffffff;
          padding: 12px 24px;
          display: flex;
          align-items: center;
          justify-content: space-between;
          font-family: 'Inter', sans-serif;
          box-shadow: 0 2px 8px rgba(0,0,0,0.3);
          position: sticky;
          top: 0;
          z-index: 1000;
        }}
        .header-bar h1 {{
          font-size: 1.1rem;
          font-weight: 600;
          margin: 0;
        }}
        .header-bar .meta {{
          font-size: 0.85rem;
          color: #94a3b8;
        }}
        .pdf-container {{
          display: flex;
          flex-direction: column;
          align-items: center;
          padding: 24px 0;
          gap: 24px;
        }}
        .pdf-page {{
          background: var(--color-bg-page);
          width: 595.0pt;
          min-height: 842.0pt;
          position: relative;
          overflow: visible;
          box-shadow: var(--page-shadow);
          page-break-after: always;
          border-radius: 2px;
        }}
        .pdf-page p {{
          position: absolute;
          margin: 0;
          padding: 0;
          white-space: pre-wrap;
          word-break: break-word;
          line-height: 1.15;
          z-index: 10;
        }}
        .vector-fill-box {{
          position: absolute;
          pointer-events: none;
          z-index: 2;
          box-sizing: border-box;
        }}
        .vector-box {{
          position: absolute;
          pointer-events: none;
          z-index: 3;
          box-sizing: border-box;
        }}
        .img-box {{
          position: absolute;
          z-index: 5;
          object-fit: contain;
        }}
        @media print {{
          .header-bar {{ display: none; }}
          body {{ background: #ffffff; }}
          .pdf-container {{ padding: 0; gap: 0; }}
          .pdf-page {{ box-shadow: none; margin: 0; border-radius: 0; }}
        }}
        """,
        "</style>",
        "</head>",
        "<body>",
        f"<div class='header-bar'>",
        f"  <h1>Oncquest Report Renderer — {doc_title}</h1>",
        f"  <div class='meta'>Source: {sanitize_text(source_file)} | Pages: {len(pages)}</div>",
        f"</div>",
        "<div class='pdf-container'>"
    ]

    for p in pages:
        p_num = p.get("page_number", 1)
        pw = p.get("width", 595.0)
        ph = p.get("height", 842.0)
        hy_cutoff = p.get("header_y_cutoff", 0.0)
        fy_cutoff = p.get("footer_y_cutoff", ph)

        html_parts.append(f"<div class='pdf-page' id='page-{p_num}' style='width:{pw:.1f}pt; min-height:{ph:.1f}pt;'>")

        # 1. Render Vector Drawings in body region ONLY
        for d in p.get("drawings", []):
            bbox = d.get("bbox")
            if not bbox or len(bbox) < 4:
                continue
            x0, y0, x1, y1 = bbox
            if y0 < hy_cutoff or y1 > fy_cutoff:
                continue
            w = max(0.5, x1 - x0)
            h = max(0.5, y1 - y0)
            fill_col = d.get("fill_color")
            stroke_col = d.get("stroke_color")

            if fill_col:
                html_parts.append(
                    f"<div class='vector-fill-box' style='left:{x0:.2f}pt; top:{y0:.2f}pt; width:{w:.2f}pt; height:{h:.2f}pt; background-color:{fill_col};'></div>"
                )
            elif stroke_col:
                stroke_w = d.get("width") or 1.0
                html_parts.append(
                    f"<div class='vector-box' style='left:{x0:.2f}pt; top:{y0:.2f}pt; width:{w:.2f}pt; height:{h:.2f}pt; border:{stroke_w}px solid {stroke_col};'></div>"
                )

        # 2. Render Images in body region ONLY
        for img in p.get("images", []):
            bbox = img.get("bbox")
            data_uri = img.get("data_uri")
            if bbox and data_uri:
                x0, y0, x1, y1 = bbox
                if y0 < hy_cutoff or y1 > fy_cutoff:
                    continue
                w = max(1.0, x1 - x0)
                h = max(1.0, y1 - y0)
                html_parts.append(
                    f"<img class='img-box' src='{data_uri}' style='left:{x0:.2f}pt; top:{y0:.2f}pt; width:{w:.2f}pt; height:{h:.2f}pt;' />"
                )

        # 3. Render Text Blocks in body region ONLY
        for b in p.get("text_blocks", []):
            b_bbox = b.get("bbox", [0, 0, 0, 0])
            if b_bbox[1] < hy_cutoff or b_bbox[3] > fy_cutoff:
                continue

            lines = b.get("lines")
            if lines:
                for line in lines:
                    l_bbox = line.get("bbox", [0, 0, 0, 0])
                    lx0, ly0, lx1, ly1 = l_bbox
                    if ly0 < hy_cutoff or ly1 > fy_cutoff:
                        continue
                    spans = line.get("spans", [])

                    if not spans:
                        continue

                    spans_html = []
                    for s in spans:
                        raw_text = s.get("text", "")
                        if not raw_text:
                            continue
                        clean_t = sanitize_text(raw_text)
                        font_name = s.get("font", "")
                        font_family = get_font_family(font_name)
                        size = s.get("size", 10.0)
                        color = s.get("color", "#000000")
                        is_bold = s.get("bold") or ("bold" in font_name.lower())
                        is_italic = s.get("italic") or ("italic" in font_name.lower())

                        font_wt = "bold" if is_bold else "normal"
                        font_st = "italic" if is_italic else "normal"

                        span_style = (
                            f"font-size:{size:.2f}pt; "
                            f"color:{color}; "
                            f"font-family:{font_family}; "
                            f"font-weight:{font_wt}; "
                            f"font-style:{font_st};"
                        )
                        spans_html.append(f"<span style='{span_style}'>{clean_t}</span>")

                    if spans_html:
                        html_parts.append(
                            f"<p style='left:{lx0:.2f}pt; top:{ly0:.2f}pt;'>{''.join(spans_html)}</p>"
                        )
            else:
                raw_text = b.get("text", "")
                if not raw_text:
                    continue
                clean_t = sanitize_text(raw_text)
                font_name = b.get("font", "")
                font_family = get_font_family(font_name)
                size = b.get("max_font_size") or b.get("size") or 10.0
                color = b.get("color", "#000000")
                is_bold = b.get("is_bold") or b.get("bold") or ("bold" in font_name.lower())
                is_italic = b.get("is_italic") or b.get("italic") or ("italic" in font_name.lower())

                font_wt = "bold" if is_bold else "normal"
                font_st = "italic" if is_italic else "normal"

                span_style = (
                    f"font-size:{size:.2f}pt; "
                    f"color:{color}; "
                    f"font-family:{font_family}; "
                    f"font-weight:{font_wt}; "
                    f"font-style:{font_st};"
                )
                lx0, ly0 = b_bbox[0], b_bbox[1]
                html_parts.append(
                    f"<p style='left:{lx0:.2f}pt; top:{ly0:.2f}pt; {span_style}'>{clean_t}</p>"
                )

        html_parts.append("</div>")

    html_parts.append("</div></body></html>")
    full_html = "\n".join(html_parts)

    if output_path:
        out_p = Path(output_path)
        out_p.parent.mkdir(parents=True, exist_ok=True)
        with open(out_p, "w", encoding="utf-8") as f_out:
            f_out.write(full_html)

    return full_html


def convert_html_to_docx(html_input, output_path: str = None, theme_config: dict = None):
    """
    Converts an HTML string or HTML file path directly into a Microsoft Word (.docx) file.
    Decouples HTML rendering from Word export so users can preview/edit HTML first,
    and then trigger HTML -> Word conversion on demand.
    Returns bytes of the Word file, or writes to output_path if provided.
    """
    import tempfile

    # Try high-fidelity PDF -> Word conversion using Playwright and pdf2docx
    if sync_playwright is not None and PDF2DocxConverter is not None:
        try:
            with tempfile.TemporaryDirectory() as tmp_dir:
                tmp_dir_path = Path(tmp_dir)
                html_file_path = tmp_dir_path / "temp.html"

                if isinstance(html_input, (str, Path)) and os.path.exists(str(html_input)) and os.path.isfile(str(html_input)):
                    with open(str(html_input), "r", encoding="utf-8") as f:
                        html_content = f.read()
                else:
                    html_content = str(html_input)

                html_file_path.write_text(html_content, encoding="utf-8")

                compiled_pdf = tmp_dir_path / "compiled.pdf"
                render_html_to_pdf_and_preview(html_file_path, compiled_pdf)

                if compiled_pdf.exists() and compiled_pdf.stat().st_size > 0:
                    out_docx_path = tmp_dir_path / "output.docx"
                    success = convert_pdf_to_word(compiled_pdf, out_docx_path, theme_config=theme_config)
                    if success and out_docx_path.exists():
                        if output_path:
                            out_p = Path(output_path)
                            out_p.parent.mkdir(parents=True, exist_ok=True)
                            with open(out_docx_path, "rb") as f_in, open(out_p, "wb") as f_out:
                                f_out.write(f_in.read())
                            return None
                        else:
                            with open(out_docx_path, "rb") as f_in:
                                return f_in.read()
        except Exception as e:
            print(f"[*] Note: PDF-based HTML-to-Word conversion failed: {e}. Falling back to direct HTML parse.")

    # Original Fallback Implementation
    from docx import Document
    from docx.shared import Inches
    from bs4 import BeautifulSoup
    try:
        from htmldocx import HtmlToDocx
    except ImportError:
        HtmlToDocx = None

    if isinstance(html_input, (str, Path)) and os.path.exists(str(html_input)) and os.path.isfile(str(html_input)):
        with open(str(html_input), "r", encoding="utf-8") as f:
            html_content = f.read()
    else:
        html_content = str(html_input)

    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    soup = BeautifulSoup(html_content, 'html.parser')
    for tag in soup(['style', 'script', 'meta', 'link']):
        tag.decompose()

    body = soup.find('body')
    clean_html = str(body) if body else str(soup)

    converted_with_htmldocx = False
    if HtmlToDocx is not None:
        try:
            parser = HtmlToDocx()
            parser.add_html_to_document(clean_html, doc)
            converted_with_htmldocx = True
        except Exception:
            converted_with_htmldocx = False

    if not converted_with_htmldocx:
        _parse_html_soup_to_docx(soup, doc)

    if output_path:
        out_p = Path(output_path)
        out_p.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(out_p))
        return None
    else:
        buf = io.BytesIO()
        doc.save(buf)
        return buf.getvalue()


def _parse_html_soup_to_docx(soup, doc):
    """Fallback manual parser from BeautifulSoup to python-docx."""
    for elem in soup.find_all(['h1', 'h2', 'h3', 'h4', 'p', 'table', 'ul', 'ol', 'div']):
        if elem.name in ['h1', 'h2', 'h3', 'h4']:
            level = int(elem.name[1])
            txt = elem.get_text(strip=True)
            if txt:
                doc.add_heading(txt, level=level)
        elif elem.name in ['p', 'div'] and elem.find_parent('table') is None:
            txt = elem.get_text(strip=True)
            if txt:
                doc.add_paragraph(txt)
        elif elem.name == 'table' and elem.find_parent('table') is None:
            rows = elem.find_all('tr')
            if rows:
                cols_count = max(len(r.find_all(['td', 'th'])) for r in rows)
                if cols_count > 0:
                    t = doc.add_table(rows=len(rows), cols=cols_count)
                    for r_idx, r in enumerate(rows):
                        cells = r.find_all(['td', 'th'])
                        for c_idx, c in enumerate(cells):
                            if c_idx < cols_count:
                                t.rows[r_idx].cells[c_idx].text = c.get_text(strip=True)


def convert_pdf_via_pdf2docx(pdf_path, docx_path):
    """
    Directly converts a PDF file to Word (.docx) using pdf2docx Converter.
    """
    if PDF2DocxConverter is None:
        print("[!] pdf2docx library is not installed. Install via: pip install pdf2docx")
        return False
    try:
        pdf_p = str(Path(pdf_path).absolute())
        docx_p = str(Path(docx_path).absolute())
        Path(docx_p).parent.mkdir(parents=True, exist_ok=True)

        cv_obj = PDF2DocxConverter(pdf_p)
        cv_obj.convert(docx_p)
        cv_obj.close()
        print(f"   [+] pdf2docx conversion completed successfully: {docx_p}")
        return True
    except Exception as e:
        print(f"   [!] pdf2docx conversion error for {pdf_path}: {e}")
        return False


def convert_pdf_full_pipeline(pdf_path, output_dir=None, theme_config: dict = None):
    """
    Executes full 4-step pipeline:
    1. Input PDF -> Extract JSON
    2. JSON -> Render HTML template
    3. HTML -> Compile Intermediate PDF (via Playwright)
    4. Intermediate PDF -> Word (.docx) via pdf2docx (fallback to python-docx)
    """
    pdf_path = Path(pdf_path).absolute()
    if output_dir is None:
        output_dir = pdf_path.parent / "output"
    else:
        output_dir = Path(output_dir).absolute()
    output_dir.mkdir(parents=True, exist_ok=True)
    json_dir = pdf_path.parent / "extracted_jsons"
    json_dir.mkdir(parents=True, exist_ok=True)

    stem = pdf_path.stem

    print(f"\n==================================================")
    print(f"[*] Executing 4-Step Pipeline for: {pdf_path.name}")
    print(f"==================================================")

    # Step 1: PDF -> JSON
    print(f"\n[Step 1/4] Extracting PDF to JSON...")
    extracted_data = extract_report_data(str(pdf_path))
    json_path = json_dir / f"{stem}.json"
    if extracted_data:
        with open(json_path, "w", encoding="utf-8") as f_json:
            json.dump(extracted_data, f_json, indent=2, ensure_ascii=False)
        print(f"   [+] JSON saved: {json_path}")

    # Step 2: JSON -> HTML
    print(f"\n[Step 2/4] Rendering JSON to HTML...")
    out_html = output_dir / f"{stem}_template.html"
    doc = fitz.open(str(pdf_path))
    full_html = render_exact_pdf_layout_html(doc, doc_title=pdf_path.name, theme_config=theme_config)
    doc.close()

    with open(out_html, "w", encoding="utf-8") as f_html:
        f_html.write(full_html)
    print(f"   [+] HTML saved: {out_html}")

    # Step 3: HTML -> Compiled Result PDF
    print(f"\n[Step 3/4] Compiling HTML to Result PDF...")
    intermediate_pdf = output_dir / f"{stem}_compiled.pdf"
    render_html_to_pdf_and_preview(out_html, intermediate_pdf)
    print(f"   [+] Result PDF compiled from HTML: {intermediate_pdf}")

    # Step 4: Convert Compiled Result PDF to Word (.docx) using pdf2docx
    print(f"\n[Step 4/4] Converting Compiled Result PDF to Word (.docx) via pdf2docx...")
    out_docx = output_dir / f"{stem}_report.docx"
    target_pdf_for_word = intermediate_pdf if (intermediate_pdf.exists() and intermediate_pdf.stat().st_size > 0) else pdf_path
    convert_pdf_to_word(target_pdf_for_word, out_docx)

    print(f"\n[+] 4-Step Pipeline Completed Successfully! Final Word doc: {out_docx}")
    print(f"==================================================\n")
    return out_docx


def convert_pdf_to_word(pdf_path, docx_path, theme_config: dict = None):
    """
    Convert PDF to Word (.docx) directly using pdf2docx Converter (pdftoDoc logic).
    """
    pdf_p = str(Path(pdf_path).absolute())
    docx_p = str(Path(docx_path).absolute())
    Path(docx_p).parent.mkdir(parents=True, exist_ok=True)

    if PDF2DocxConverter is None:
        print("[!] Error: pdf2docx is not installed. Run 'pip install pdf2docx'")
        return False

    try:
        cv_obj = PDF2DocxConverter(pdf_p)
        cv_obj.convert(docx_p)
        cv_obj.close()
        print(f"   [+] File Converted Successfully: {docx_p}")
        return True
    except Exception as e:
        print(f"   [!] Conversion Failed: {e}")
        return False



def _get_page_bounds(page, fallback_left, fallback_width):
    """
    Compute actual content bounds from table positions.
    """
    tabs = page.find_tables()
    bboxes = []
    for tab in tabs.tables:
        if hasattr(tab, 'bbox'):
            x0, y0, x1, y1 = tab.bbox
            if y1 < 750 and (x1 - x0) > 80:
                bboxes.append(tab.bbox)

    if bboxes:
        page_left = min(b[0] for b in bboxes)
        page_right = max(b[2] for b in bboxes)
        page_width = page_right - page_left
        return page_left, page_width

    return fallback_left, fallback_width


def get_page_section_overlays(page, page_left_str, page_width_str, hy_cutoff: float = 0.0, fy_cutoff: float = 9999.0):
    """
    Extract section bounding boxes based on bold section headings in body region ONLY.
    """
    return []
    spans = []
    blocks = page.get_text('dict')['blocks']
    for b in blocks:
        if 'lines' in b:
            for ln in b['lines']:
                for s in ln['spans']:
                    if s.get('text', '').strip():
                        spans.append(s)
    spans.sort(key=lambda s: (round(s['bbox'][1], 1), round(s['bbox'][0], 1)))

    overlays = []
    for i, s in enumerate(spans):
        text = s['text'].strip()
        s_y0 = s['bbox'][1]
        s_y1 = s['bbox'][3]
        if s_y0 < hy_cutoff or s_y1 > fy_cutoff:
            continue

        is_bold = s.get('flags', 0) & 2 or "bold" in s.get('font', '').lower()
        if is_bold and text.endswith(':') and len(text) < 40:
            h_bbox = s['bbox']
            content_spans = []
            for j in range(i + 1, len(spans)):
                next_text = spans[j]['text'].strip()
                next_y0 = spans[j]['bbox'][1]
                if next_y0 > fy_cutoff:
                    break
                next_is_bold = spans[j].get('flags', 0) & 2 or "bold" in spans[j].get('font', '').lower()
                if (next_is_bold and next_text.endswith(':')) or spans[j]['bbox'][1] - h_bbox[1] > 100:
                    break
                content_spans.append(spans[j])

            if content_spans:
                c_min_y = max(hy_cutoff, h_bbox[1] - 2.0)
                c_max_y = min(fy_cutoff, max(cs['bbox'][3] for cs in content_spans) + 4.0)
                h = max(14.0, c_max_y - c_min_y)
                if c_min_y >= hy_cutoff and c_max_y <= fy_cutoff:
                    overlays.append(
                        f"<div class='section-content-box' "
                        f"style='left:{page_left_str};top:{c_min_y:.1f}pt;"
                        f"width:{page_width_str};height:{h:.1f}pt;'></div>"
                    )
    return overlays



def process_pdf(pdf_input, output_html_path=None, is_target=False, use_template=False, theme_config=None, save_output=False):
    """
    Universal PDF conversion function.
    Extracts structured JSON data and renders HTML.
    By default (save_output=False), does NOT write any files to the output directory.
    """
    doc_title = "Uploaded Report"
    if isinstance(pdf_input, (str, Path)):
        tmp_path = str(pdf_input)
        doc_title = Path(pdf_input).name
        doc = fitz.open(tmp_path)
    elif isinstance(pdf_input, bytes):
        doc = fitz.open(stream=pdf_input, filetype="pdf")
        tmp_path = None
    else:
        pdf_bytes = pdf_input.read()
        if hasattr(pdf_input, "seek"):
            pdf_input.seek(0)
        doc_title = getattr(pdf_input, "name", "Uploaded Report")
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        tmp_path = None

    extracted_data = None
    if tmp_path and os.path.exists(tmp_path):
        try:
            extracted_data = extract_report_data(tmp_path)
        except Exception:
            pass

    # 3. HTML Rendering based on requested mode
    if use_template and extracted_data:
        full_html = generate_dynamic_template_html(extracted_data, doc_title=doc_title, theme_config=theme_config)
    else:
        # Preserve 100% exact 1-to-1 visual coordinates from input PDF
        full_html = render_exact_pdf_layout_html(doc, doc_title=doc_title, theme_config=theme_config)

    # Save to disk ONLY if save_output is explicitly requested
    if save_output and output_html_path:
        output_html_path = Path(output_html_path)
        output_html_path.parent.mkdir(parents=True, exist_ok=True)
        with open(output_html_path, "w", encoding="utf-8") as f:
            f.write(full_html)

    doc.close()
    return full_html


def render_html_to_pdf_and_preview(html_path, output_pdf_path, preview_img_path=None):
    """
    Render an HTML file to PDF via Playwright Chromium (if installed).
    """
    html_path = Path(html_path).absolute()
    output_pdf_path = Path(output_pdf_path)
    output_pdf_path.parent.mkdir(parents=True, exist_ok=True)

    if sync_playwright is None:
        return output_pdf_path

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch()
            page = browser.new_page(viewport={"width": 1000, "height": 1200})
            page.goto(html_path.as_uri(), wait_until="networkidle")

            if preview_img_path:
                preview_img_path = Path(preview_img_path)
                preview_img_path.parent.mkdir(parents=True, exist_ok=True)
                page.screenshot(path=str(preview_img_path), full_page=False)

            page.pdf(
                path=str(output_pdf_path),
                print_background=True,
                prefer_css_page_size=True,
                margin={"top": "0px", "right": "0px", "bottom": "0px", "left": "0px"},
            )
            page.close()
            browser.close()
    except Exception as err:
        pass

    return output_pdf_path
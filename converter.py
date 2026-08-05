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

from extractor import extract_report_data


def render_exact_pdf_layout_html(doc, doc_title: str = "Uploaded Document", theme_config: dict = None) -> str:
    """
    Renders an HTML document where EVERY SINGLE ELEMENT (text, headers, key-values, tables,
    section boxes, images, graphs) stays in its EXACT 1-to-1 position as in the input PDF document.
    Applies user-selected theme typography, primary colors, table header styling, and cell borders.
    """
    if not theme_config:
        theme_config = {}

    primary_color = theme_config.get("primary_color", "#1f497d")
    bg_page = theme_config.get("bg_page", "#ffffff")
    text_color = theme_config.get("text_color", "#0d0d0d")
    border_color = theme_config.get("border_color", primary_color)
    font_family = theme_config.get("font_family", "Cambria, 'Times New Roman', serif")
    banner_font_min_pt = theme_config.get("banner_font_size_pt", 12.5)

    fallback_left = 35.5
    fallback_width = 524.0

    html_parts = [
        "<!DOCTYPE html>",
        "<html>",
        "<head>",
        "<meta charset='utf-8'>",
        f"<title>{doc_title}</title>",
        "<style>",
        "@page { size: 595.6pt 842.0pt; margin: 0; }",
        "* { box-sizing: border-box; }",
        f"body {{ margin: 0; padding: 0; background-color: #525659; font-family: {font_family}; color: {text_color}; }}",
        ".pdf-container { display: flex; flex-direction: column; align-items: center; padding: 20px 0; }",
        f".pdf-page {{ background: {bg_page}; width: 595.6pt; min-height: 842.0pt; margin-bottom: 20px; position: relative; overflow: visible; box-shadow: 0 4px 12px rgba(0,0,0,0.3); page-break-after: always; font-family: {font_family}; }}",
        f"div[id^='page'] {{ position: relative !important; width: 595.6pt !important; min-height: 842.0pt !important; overflow: visible !important; }}",
        f"div[id^='page'] p {{ position: absolute !important; margin: 0 !important; padding: 0 !important; white-space: normal !important; word-break: normal !important; overflow-wrap: break-word !important; max-width: 100% !important; z-index: 10 !important; font-family: {font_family} !important; line-height: 1.3 !important; overflow: visible !important; }}",
        f"div[id^='page'] span {{ word-break: normal !important; overflow-wrap: break-word !important; }}",
        f".black-banner-span {{ color: #ffffff !important; display: inline-block !important; text-align: center !important; padding: 4px 0 !important; font-weight: bold !important; font-size: 13.0pt !important; font-family: {font_family} !important; border-radius: 0px !important; white-space: nowrap !important; }}",
        f".label-bar-span {{ color: #ffffff !important; display: inline-block !important; padding: 2px 6px !important; border-radius: 2px !important; font-family: {font_family} !important; word-break: break-word !important; }}",
        f".table-header-cell {{ position: absolute; background-color: {primary_color} !important; color: #ffffff !important; font-family: {font_family} !important; font-size: 9.5pt !important; font-weight: bold !important; display: flex !important; align-items: center !important; justify-content: center !important; text-align: center !important; padding: 2px 4px !important; white-space: normal !important; word-break: break-word !important; overflow-wrap: break-word !important; line-height: 1.15 !important; border: 1px solid {border_color} !important; box-sizing: border-box !important; z-index: 15 !important; }}",
        "div[id^='page'] img { position: absolute !important; transform-origin: 0 0 !important; z-index: 5 !important; opacity: 1 !important; visibility: visible !important; display: inline-block !important; }",
        f".table-grid-cell {{ position: absolute; border: 1px solid {border_color} !important; background: transparent; pointer-events: none; z-index: 4; }}",
        f".section-content-box {{ position: absolute; border: 1px solid {border_color} !important; background: transparent; pointer-events: none; z-index: 4; border-radius: 3px; }}",
        "@media print { body { background-color: #ffffff; } .pdf-container { padding: 0; } .pdf-page { margin: 0; box-shadow: none; } }",
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
        page_right_val = page_left_val + page_width_val

        html_parts.append("<div class='pdf-page'>")
        page_html = page.get_text("html")

        # Extract PyMuPDF table headers & grid coordinates for exact alignment
        tabs = page.find_tables()
        table_header_html_divs = []
        table_grid_html_divs = []
        header_y_ranges = []

        for tab in tabs.tables:
            valid_cells = [c for c in tab.cells if c]
            if valid_cells:
                min_y0 = min(c[1] for c in valid_cells)
                header_cells = [c for c in valid_cells if abs(c[1] - min_y0) < 3.0]
                header_cells.sort(key=lambda c: c[0])
                if len(header_cells) >= 2:
                    hy0 = min(c[1] for c in header_cells)
                    hy1 = max(c[3] for c in header_cells)
                    header_y_ranges.append((hy0 - 1.0, hy1 + 1.0))

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

        cleaned = page_html
        cleaned = re.sub(r'font-family:[^;"]+', f'font-family: {font_family}', cleaned)

        # Image matrix positioning fit
        def fit_page_imgs(page_html_str):
            img_matches = list(re.finditer(r'<img\s+([^>]*style=["\']([^"\']+)["\'][^>]*)>', page_html_str))
            if not img_matches:
                return page_html_str

            img_infos = []
            for m in img_matches:
                full_tag = m.group(0)
                style_str = m.group(2)
                src_m = re.search(r'src=["\']data:image/png;base64,([^"\']+)["\']', full_tag)
                b64_data = src_m.group(1) if src_m else None

                pw, ph = 0, 0
                if b64_data:
                    try:
                        im = Image.open(io.BytesIO(base64.b64decode(b64_data)))
                        pw, ph = im.size
                    except Exception:
                        pass

                matrix_m = re.search(r'matrix\(([^)]+)\)', style_str)
                if matrix_m:
                    parts = [float(x.strip()) for x in matrix_m.group(1).split(',')]
                    sx, sy, tx, ty = parts[0], parts[3], parts[4], parts[5]
                else:
                    sx, sy, tx, ty = 1.0, 1.0, page_left_val, 0.0

                img_infos.append({
                    'tag': full_tag, 'style': style_str,
                    'pw': pw, 'ph': ph,
                    'sx': sx, 'sy': sy, 'tx': tx, 'ty': ty
                })

            valid_imgs = [info for info in img_infos if (-100 <= info['tx'] <= 595.6)]
            for info in img_infos:
                if info not in valid_imgs:
                    page_html_str = page_html_str.replace(info['tag'], '')

            for info in valid_imgs:
                new_tx = max(page_left_val, info['tx']) if info['tx'] < page_left_val else info['tx']
                pw = info['pw'] if info['pw'] > 0 else 500
                rw = pw * info['sx']
                new_sx = info['sx']
                new_sy = info['sy']
                if new_tx + rw > page_right_val:
                    avail_w = page_right_val - new_tx
                    new_sx = avail_w / pw
                    new_sy = new_sx
                if new_tx != info['tx'] or new_sx != info['sx']:
                    new_matrix = f"matrix({new_sx:.6f},0,0,{new_sy:.6f},{new_tx:.2f},{info['ty']:.2f})"
                    new_style = re.sub(r'matrix\([^)]+\)', new_matrix, info['style'])
                    new_tag = info['tag'].replace(info['style'], new_style)
                    page_html_str = page_html_str.replace(info['tag'], new_tag)

            return page_html_str

        cleaned = fit_page_imgs(cleaned)

        # Headings format
        def format_heading_p(match):
            p_tag = match.group(0)
            size_m = re.search(r'font-size:\s*([\d.]+)pt', p_tag)
            is_banner_size = bool(size_m) and float(size_m.group(1)) >= banner_font_min_pt
            if "color:#ffffff" in p_tag and is_banner_size:
                top_m = re.search(r'top:([\d\.]+)pt', p_tag)
                top_val = top_m.group(1) if top_m else "100.0"
                text_val = re.sub(r'<[^>]+>', '', p_tag).strip()
                return (
                    f'<p style="top:{top_val}pt;left:{page_left_str};'
                    f'width:{page_width_str};margin:0;padding:0;z-index:12;">'
                    f'<span class="black-banner-span" '
                    f'style="width:{page_width_str};background-color:#404040;">'
                    f'{text_val}</span></p>'
                )
            return p_tag

        cleaned = re.sub(r'<p\s+[^>]*>.*?</p>', format_heading_p, cleaned, flags=re.DOTALL)

        # Label bars format
        def format_labelbar_p(match):
            p_tag = match.group(0)
            if ("color:#ffffff" in p_tag and "black-banner-span" not in p_tag and "table-header-cell" not in p_tag):
                top_m = re.search(r'top:([\d\.]+)pt', p_tag)
                if top_m:
                    p_tag_mod = re.sub(r'left:[\d\.]+pt', f'left:{page_left_str}', p_tag)
                    if 'left:' not in p_tag:
                        p_tag_mod = p_tag_mod.replace('style="', f'style="left:{page_left_str};')
                    p_tag_mod = re.sub(r'width:[\d\.]+pt', f'width:{page_width_str}', p_tag_mod)

                    def fix_span(sm):
                        span = sm.group(0)
                        span = re.sub(r'background-color:\s*[^;"]+', f'background-color:{primary_color}', span)
                        if 'background-color' not in span:
                            span = span.replace('style="', f'style="background-color:{primary_color};')
                        span = re.sub(r'width:\s*[^;"]+', f'width:{page_width_str}', span)
                        if 'width:' not in span:
                            span = span.replace('style="', f'style="width:{page_width_str};')
                        if 'display:' not in span:
                            span = span.replace('style="', 'style="display:inline-block;')
                        return span

                    return re.sub(r'<span\s+[^>]*>.*?</span>', fix_span, p_tag_mod, flags=re.DOTALL)
            return p_tag

        cleaned = re.sub(r'<p\s+[^>]*>.*?</p>', format_labelbar_p, cleaned, flags=re.DOTALL)

        section_overlays = get_page_section_overlays(page, page_left_str, page_width_str)

        html_parts.append(cleaned)
        html_parts.extend(section_overlays)
        html_parts.extend(table_header_html_divs)
        html_parts.extend(table_grid_html_divs)
        html_parts.append("</div>")

    html_parts.append("</div></body></html>")
    return "\n".join(html_parts)


def generate_dynamic_template_html(data: dict, doc_title: str = "Uploaded Document", theme_config: dict = None) -> str:
    """
    Renders extracted PDF JSON data dynamically into a customizable HTML template with clean section boxes,
    structured key-value cards, responsive tables, and zero sentence cut-offs.
    """
    if not theme_config:
        theme_config = {}

    primary_color = theme_config.get("primary_color", "#1f497d")
    secondary_color = theme_config.get("secondary_color", "#008080")
    bg_page = theme_config.get("bg_page", "#ffffff")
    text_color = theme_config.get("text_color", "#0d0d0d")
    border_color = theme_config.get("border_color", primary_color)
    table_header_bg = theme_config.get("table_header_bg", primary_color)
    table_header_text = theme_config.get("table_header_text", "#ffffff")
    font_family = theme_config.get("font_family", "Cambria, 'Times New Roman', serif")
    header_title = theme_config.get("header_title", doc_title.upper().replace(".PDF", ""))
    header_subtitle = theme_config.get("header_subtitle", "Universal Dynamic Document Report")

    show_kv = theme_config.get("show_kv", True)
    show_tables = theme_config.get("show_tables", True)
    show_sections = theme_config.get("show_sections", True)
    show_images = theme_config.get("show_images", True)

    show_badges = theme_config.get("show_badges", True)
    badge_rules = theme_config.get("badge_rules", {
        "danger": ["pathogenic", "positive", "high", "failed", "rejected", "overdue",
                   "invalid", "expired", "critical", "denied", "delinquent", "abnormal"],
        "warning": ["vus", "uncertain", "warning", "pending", "under review",
                    "partial", "provisional", "conditional"],
        "success": ["passed", "normal", "negative", "approved", "paid", "valid",
                    "cleared", "completed", "compliant", "settled"]
    })

    show_footer_signatures = theme_config.get("show_footer_signatures", True)
    footer_signature_labels = theme_config.get(
        "footer_signature_labels",
        ["Prepared / Verified By", "Reviewing Officer", "Authorized Signatory"]
    )

    kv = data.get("all_key_value_pairs") or data.get("extracted_key_value_pairs") or {}
    tables = data.get("all_tables") or data.get("tables") or []
    sections = data.get("all_boxes_and_sections") or data.get("content_sections") or []
    images = data.get("all_images_and_graphs") or data.get("images_and_graphs") or []
    summary = data.get("document_summary") or {}

    # 1. Header Logo & Branding
    logo_html = ""
    header_images = [img for img in images if img.get("page") == 1 and img.get("width", 0) > 40]
    if header_images and show_images:
        first_img = header_images[0]
        logo_html = f'<img src="{first_img.get("data_uri")}" style="max-height: 55px; max-width: 220px; object-fit: contain;" alt="Logo" />'

    # 2. Key-Value Pairs Grid HTML
    header_rows_html = ""
    if show_kv and kv:
        kv_items = list(kv.items())
        for i in range(0, len(kv_items), 2):
            k1, v1 = kv_items[i]
            k2, v2 = kv_items[i + 1] if (i + 1) < len(kv_items) else ("", "")

            td_k2 = f'<td class="header-label">{k2}:</td><td class="header-val">{v2}</td>' if k2 else '<td class="header-label"></td><td class="header-val"></td>'
            header_rows_html += f"""
            <tr>
                <td class="header-label">{k1}:</td>
                <td class="header-val"><strong>{v1}</strong></td>
                {td_k2}
            </tr>
            """

    if not header_rows_html:
        header_rows_html = f"""
        <tr>
            <td class="header-label">Document Name:</td>
            <td class="header-val"><strong>{summary.get("file_name", doc_title)}</strong></td>
            <td class="header-label">Total Pages:</td>
            <td class="header-val">{summary.get("total_pages", 1)}</td>
        </tr>
        """

    # 3. Content Section Boxes HTML (clean boxed layout)
    sections_html = ""
    if show_sections and sections:
        for sec in sections:
            title = sec.get("title", "").strip()
            sec_type = sec.get("type", "")
            if sec_type == "demographics_box" or title.startswith("Header & Metadata Box"):
                continue

            content_text = sec.get("content_text", [])
            if isinstance(content_text, list):
                body_paragraphs = "".join([f'<p class="section-p">{t.strip()}</p>' for t in content_text if t and t.strip()])
            else:
                body_paragraphs = f'<p class="section-p">{str(content_text).strip()}</p>'

            if not title and not body_paragraphs:
                continue

            sec_title_html = f'<div class="section-title">{title}</div>' if title else ''
            sections_html += f"""
            <div class="section-box">
                {sec_title_html}
                <div class="section-body">
                    {body_paragraphs}
                </div>
            </div>
            """

    # 4. Tables HTML
    tables_html = ""
    if show_tables and tables:
        for t_idx, tab in enumerate(tables):
            headers = tab.get("headers", [])
            rows = tab.get("rows", [])
            page_n = tab.get("page", 1)
            if not headers and not rows:
                continue

            th_html = "".join([f"<th>{h}</th>" for h in headers]) if headers else ""
            tr_html = ""
            for r in rows:
                tds = ""
                for cell in r:
                    cell_str = str(cell).replace('\n', '<br>')
                    cell_lower = cell_str.lower()
                    cell_formatted = cell_str
                    if show_badges:
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
                <div class="section-title">Extracted Images & Graphical Content</div>
                <div class="image-grid">
                    {img_cards}
                </div>
            </div>
            """

    # 6. Footer Signatures HTML
    footer_html = ""
    if show_footer_signatures and footer_signature_labels:
        sig_boxes = "".join(
            f"""
            <div class="sig-box">
                <div>______________________</div>
                <div class="sig-title">{label}</div>
            </div>
            """
            for label in footer_signature_labels
        )
        footer_html = f'<div class="footer-signatures">{sig_boxes}</div>'

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
.pdf-page {{ background: {bg_page}; width: 595.6pt; min-height: 842.0pt; padding: 35.5pt; margin-bottom: 20px; position: relative; box-shadow: 0 4px 12px rgba(0,0,0,0.15); page-break-after: always; font-family: {font_family}; word-break: normal; overflow-wrap: break-word; }}
.logo-header {{ display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid {primary_color}; padding-bottom: 8pt; margin-bottom: 12pt; }}
.brand-title {{ font-size: 18pt; font-weight: bold; color: {primary_color}; letter-spacing: 0.02em; }}
.brand-subtitle {{ font-size: 9pt; color: #64748b; margin-top: 2pt; }}
.header-table {{ width: 100%; border-collapse: collapse; margin-bottom: 14pt; border: 1px solid {border_color}; border-radius: 4px; overflow: hidden; table-layout: fixed; }}
.header-table td {{ padding: 6pt 8pt; font-size: 9.5pt; color: {text_color}; border: 1px solid #cbd5e1; vertical-align: middle; word-break: normal; overflow-wrap: break-word; }}
.header-label {{ font-weight: bold; color: {primary_color}; width: 22%; background-color: #f8fafc; }}
.header-val {{ width: 28%; }}
.banner-dark {{ background-color: {primary_color}; color: #ffffff; text-align: center; padding: 7pt 0; font-weight: bold; font-size: 12.5pt; font-family: {font_family}; margin: 14pt 0; text-transform: uppercase; letter-spacing: 0.04em; border-radius: 4px; box-shadow: 0 2px 4px rgba(0,0,0,0.08); }}
.section-box {{ border: 1px solid {border_color}; border-radius: 6px; margin-bottom: 14pt; position: relative; background: #ffffff; overflow: hidden; box-shadow: 0 2px 5px rgba(0,0,0,0.04); }}
.section-title {{ background-color: {primary_color}; color: #ffffff; font-weight: bold; font-size: 10.0pt; padding: 6pt 10pt; display: block; border-top-left-radius: 4px; border-top-right-radius: 4px; letter-spacing: 0.02em; }}
.section-body {{ padding: 10pt 12pt; font-size: 9.5pt; line-height: 1.5; color: {text_color}; }}
.section-p {{ margin: 0 0 6pt 0; text-align: justify; word-break: normal; overflow-wrap: break-word; }}
.section-p:last-child {{ margin-bottom: 0; }}
.table-card-box {{ margin-top: 14pt; margin-bottom: 14pt; border: 1px solid {border_color}; border-radius: 6px; overflow: hidden; background: #ffffff; box-shadow: 0 2px 5px rgba(0,0,0,0.04); }}
.table-card-header {{ font-size: 9.5pt; font-weight: bold; color: {primary_color}; padding: 6pt 10pt; background-color: #f8fafc; border-bottom: 1px solid {border_color}; }}
.table-custom {{ width: 100%; border-collapse: collapse; font-size: 9.5pt; table-layout: auto; }}
.table-custom th {{ background-color: {table_header_bg}; color: {table_header_text}; font-family: {font_family}; font-size: 9.5pt; font-weight: bold; text-align: center; padding: 7pt 6pt; border: 1px solid {border_color}; word-break: normal; overflow-wrap: break-word; }}
.table-custom td {{ padding: 6pt 8pt; border: 1px solid {border_color}; vertical-align: top; line-height: 1.35; font-size: 9.5pt; word-break: normal; overflow-wrap: break-word; }}
.badge-danger {{ background: #dc2626; color: #ffffff; padding: 2px 6px; border-radius: 3px; font-weight: bold; display: inline-block; font-size: 8.5pt; }}
.badge-warning {{ background: #d97706; color: #ffffff; padding: 2px 6px; border-radius: 3px; font-weight: bold; display: inline-block; font-size: 8.5pt; }}
.badge-success {{ background: #16a34a; color: #ffffff; padding: 2px 6px; border-radius: 3px; font-weight: bold; display: inline-block; font-size: 8.5pt; }}
.image-grid {{ display: flex; flex-wrap: wrap; gap: 12px; margin-top: 8pt; padding: 10pt; }}
.image-card {{ border: 1px solid #e2e8f0; border-radius: 4px; padding: 8px; background: #fafafa; text-align: center; max-width: 240px; }}
.extracted-img {{ max-width: 100%; max-height: 160px; object-fit: contain; border-radius: 2px; }}
.image-caption {{ font-size: 8pt; color: #64748b; margin-top: 4pt; font-family: sans-serif; }}
.footer-signatures {{ display: flex; justify-content: space-between; margin-top: 30pt; padding-top: 10pt; border-top: 1px solid #cbd5e1; font-size: 9pt; }}
.sig-box {{ text-align: center; width: 30%; }}
.sig-title {{ font-weight: bold; color: {primary_color}; margin-top: 4pt; }}
</style>
</head>
<body>
<div class="pdf-container">
  <div class="pdf-page">
    <div class="logo-header">
      <div>
        <div class="brand-title">{header_title}</div>
        <div class="brand-subtitle">{header_subtitle}</div>
      </div>
      {logo_html}
    </div>
    
    <table class="header-table">
      {header_rows_html}
    </table>

    <div class="banner-dark">{header_title}</div>

    {sections_html}

    {tables_html}

    {images_html}

    {footer_html}
  </div>
</div>
</body>
</html>"""
    return html_content


def convert_json_to_docx(data: dict, output_path: str = None, theme_config: dict = None):
    """
    Converts extracted JSON document data into a beautifully formatted Microsoft Word (.docx) file.
    Renders styled headers, metadata key-value tables, content sections with colored left callout borders,
    formatted data tables with colored header rows, images, and signature blocks.
    Returns bytes of the Word file, or writes to output_path if provided.
    """
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT

    if not theme_config:
        theme_config = {}

    primary_hex = theme_config.get("primary_color", "#1f497d").lstrip("#")
    if len(primary_hex) == 6:
        p_r, p_g, p_b = int(primary_hex[0:2], 16), int(primary_hex[2:4], 16), int(primary_hex[4:6], 16)
    else:
        p_r, p_g, p_b = 31, 73, 125
    primary_color_rgb = RGBColor(p_r, p_g, p_b)

    doc = Document()
    
    # 0.75 in margins
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    doc_summary = data.get("document_summary", {})
    file_name = doc_summary.get("file_name", "Document Report")
    header_title = theme_config.get("header_title", file_name.upper().replace(".PDF", ""))
    header_subtitle = theme_config.get("header_subtitle", "Universal Dynamic Document Report")

    # Title Banner
    p_title = doc.add_paragraph()
    run_title = p_title.add_run(header_title)
    run_title.font.name = 'Arial'
    run_title.font.size = Pt(20)
    run_title.font.bold = True
    run_title.font.color.rgb = primary_color_rgb

    p_sub = doc.add_paragraph()
    run_sub = p_sub.add_run(header_subtitle)
    run_sub.font.name = 'Arial'
    run_sub.font.size = Pt(10)
    run_sub.font.color.rgb = RGBColor(100, 116, 139)

    doc.add_paragraph()

    # 1. Metadata Key-Value Table
    kv = data.get("all_key_value_pairs") or data.get("extracted_key_value_pairs") or {}
    if kv:
        p_kv_heading = doc.add_paragraph()
        run_kvh = p_kv_heading.add_run("Header Metadata & Information")
        run_kvh.font.name = 'Arial'
        run_kvh.font.size = Pt(12)
        run_kvh.font.bold = True
        run_kvh.font.color.rgb = primary_color_rgb

        kv_items = list(kv.items())
        table_rows = (len(kv_items) + 1) // 2
        table = doc.add_table(rows=table_rows, cols=4)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        for i in range(0, len(kv_items), 2):
            r_idx = i // 2
            row = table.rows[r_idx]

            k1, v1 = kv_items[i]
            k2, v2 = kv_items[i+1] if (i+1) < len(kv_items) else ("", "")

            row.cells[0].text = f"{k1}:"
            row.cells[1].text = str(v1)
            row.cells[2].text = f"{k2}:" if k2 else ""
            row.cells[3].text = str(v2) if k2 else ""

            for col_idx in [0, 2]:
                cell = row.cells[col_idx]
                if cell.paragraphs[0].runs:
                    cell.paragraphs[0].runs[0].font.bold = True
                    cell.paragraphs[0].runs[0].font.size = Pt(9.5)
                    cell.paragraphs[0].runs[0].font.color.rgb = primary_color_rgb
            for col_idx in [1, 3]:
                cell = row.cells[col_idx]
                if cell.paragraphs[0].runs:
                    cell.paragraphs[0].runs[0].font.size = Pt(9.5)

        doc.add_paragraph()

    # 2. Content Sections
    sections = data.get("all_boxes_and_sections") or data.get("content_sections") or []
    if sections:
        for sec in sections:
            title = sec.get("title", "").strip()
            sec_type = sec.get("type", "")
            if sec_type == "demographics_box" or title.startswith("Header & Metadata Box"):
                continue

            content_text = sec.get("content_text", [])
            if isinstance(content_text, list):
                body_lines = [t.strip() for t in content_text if t.strip()]
            else:
                body_lines = [str(content_text).strip()]

            if not title and not body_lines:
                continue

            if title:
                p_sec = doc.add_paragraph()
                r_sec = p_sec.add_run(title)
                r_sec.font.name = 'Arial'
                r_sec.font.size = Pt(12)
                r_sec.font.bold = True
                r_sec.font.color.rgb = primary_color_rgb

            for line in body_lines:
                p_body = doc.add_paragraph()
                r_body = p_body.add_run(line)
                r_body.font.name = 'Arial'
                r_body.font.size = Pt(10)
                p_body.paragraph_format.line_spacing = 1.15
                p_body.paragraph_format.space_after = Pt(4)

    # 3. Data Tables
    tables = data.get("all_tables") or data.get("tables") or []
    if tables:
        for t_idx, tab in enumerate(tables):
            headers = tab.get("headers", [])
            rows = tab.get("rows", [])
            page_n = tab.get("page", 1)

            if not headers and not rows:
                continue

            p_tbl = doc.add_paragraph()
            r_tbl = p_tbl.add_run(f"Table {t_idx + 1} (Page {page_n})")
            r_tbl.font.name = 'Arial'
            r_tbl.font.size = Pt(11)
            r_tbl.font.bold = True
            r_tbl.font.color.rgb = primary_color_rgb

            total_rows = (1 if headers else 0) + len(rows)
            num_cols = max(len(headers), max((len(r) for r in rows), default=1))
            
            docx_table = doc.add_table(rows=total_rows, cols=num_cols)
            docx_table.alignment = WD_TABLE_ALIGNMENT.CENTER

            curr_r = 0
            if headers:
                hdr_cells = docx_table.rows[0].cells
                for c_i, h in enumerate(headers):
                    if c_i < len(hdr_cells):
                        hdr_cells[c_i].text = str(h)
                        if hdr_cells[c_i].paragraphs[0].runs:
                            hdr_cells[c_i].paragraphs[0].runs[0].font.bold = True
                            hdr_cells[c_i].paragraphs[0].runs[0].font.size = Pt(9.5)
                            hdr_cells[c_i].paragraphs[0].runs[0].font.color.rgb = primary_color_rgb
                curr_r += 1

            for r in rows:
                if curr_r < total_rows:
                    row_cells = docx_table.rows[curr_r].cells
                    for c_i, cell_v in enumerate(r):
                        if c_i < len(row_cells):
                            row_cells[c_i].text = str(cell_v)
                            if row_cells[c_i].paragraphs[0].runs:
                                row_cells[c_i].paragraphs[0].runs[0].font.size = Pt(9.0)
                    curr_r += 1

            doc.add_paragraph()

    if output_path:
        out_p = Path(output_path)
        out_p.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(out_p))
        return None
    else:
        buf = io.BytesIO()
        doc.save(buf)
        return buf.getvalue()


def convert_pdf_to_word(pdf_path, docx_path):
    """
    Convert a PDF file to Word (.docx) using extracted JSON and python-docx.
    """
    try:
        pdf_path = Path(pdf_path).absolute()
        docx_path = Path(docx_path).absolute()
        extracted_data = extract_report_data(str(pdf_path))
        if extracted_data:
            convert_json_to_docx(extracted_data, output_path=str(docx_path))
            return docx_path
        else:
            from pdf2docx import Converter as DocxConverter
            docx_path.parent.mkdir(parents=True, exist_ok=True)
            cv = DocxConverter(str(pdf_path))
            cv.convert(str(docx_path), start=0, end=None)
            cv.close()
            return docx_path
    except Exception as e:
        return None


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


def get_page_section_overlays(page, page_left_str, page_width_str):
    """
    Extract section bounding boxes based on bold section headings.
    """
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
        is_bold = s.get('flags', 0) & 2 or "bold" in s.get('font', '').lower()
        if is_bold and text.endswith(':') and len(text) < 40:
            h_bbox = s['bbox']
            content_spans = []
            for j in range(i + 1, len(spans)):
                next_text = spans[j]['text'].strip()
                next_is_bold = spans[j].get('flags', 0) & 2 or "bold" in spans[j].get('font', '').lower()
                if (next_is_bold and next_text.endswith(':')) or spans[j]['bbox'][1] - h_bbox[1] > 100:
                    break
                content_spans.append(spans[j])

            if content_spans:
                c_min_y = h_bbox[3] + 1.0
                c_max_y = max(cs['bbox'][3] for cs in content_spans) + 3.0
                h = max(12.0, c_max_y - c_min_y)
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
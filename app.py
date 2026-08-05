"""
Universal Dynamic PDF Extractor, JSON Inspector & Theme Renderer Web App
========================================================================
Interactive Streamlit Web UI:
1. Upload ANY PDF document format.
2. Dynamically extract JSON containing key-values, text boxes, data tables, images, and graphs.
3. Choose layout mode:
   - 📌 Exact Input PDF Layout (Preserve 100% original coordinates & positions)
   - 🏷️ Standardized HTML Flow Template (Responsive cards layout)
4. Customize design system tokens live (colors, typography, headers, component visibility).
5. Live-render HTML template, view extracted JSON tabs, and download converted HTML, JSON, and PDF files.
"""

import os
import tempfile
import base64
import json
from pathlib import Path
import streamlit as st
import fitz  # PyMuPDF
import streamlit.components.v1 as components

from converter import process_pdf, generate_dynamic_template_html, render_exact_pdf_layout_html, render_html_to_pdf_and_preview, convert_json_to_docx
from extractor import extract_report_data

# Configure Streamlit Page
st.set_page_config(
    page_title="Universal Dynamic PDF & JSON Converter",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling for Streamlit App
st.markdown("""
<style>
    .main { background-color: #f8fafc; }
    
    .header-card {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #1f497d 100%);
        color: #ffffff;
        padding: 2.0rem 2.2rem;
        border-radius: 16px;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.25);
        margin-bottom: 1.8rem;
    }
    .header-title {
        font-size: 2.1rem;
        font-weight: 800;
        margin: 0;
        letter-spacing: -0.02em;
        color: #ffffff;
    }
    .header-subtitle {
        font-size: 1.0rem;
        color: #94a3b8;
        margin-top: 0.4rem;
        margin-bottom: 0;
    }

    .content-card {
        background: #ffffff;
        border-radius: 12px;
        padding: 1.4rem;
        border: 1px solid #e2e8f0;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.04);
        margin-bottom: 1.2rem;
    }

    .metric-badge {
        display: inline-block;
        padding: 0.35em 0.8em;
        font-size: 0.85em;
        font-weight: 700;
        color: #ffffff;
        border-radius: 6px;
        background-color: #1f497d;
        margin-right: 6px;
    }

    div.stButton > button:first-child {
        background: linear-gradient(135deg, #1f497d 0%, #0f172a 100%);
        color: white;
        font-weight: 600;
        font-size: 1.05rem;
        padding: 0.65rem 1.8rem;
        border-radius: 10px;
        border: none;
        box-shadow: 0 4px 12px rgba(31, 73, 125, 0.3);
        transition: all 0.2s ease-in-out;
        width: 100%;
    }
    div.stButton > button:first-child:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 16px rgba(31, 73, 125, 0.4);
    }
</style>
""", unsafe_allow_html=True)

# Sidebar Configuration & Theme Customizer
st.sidebar.title("📐 Layout & Design Settings")

layout_mode_choice = st.sidebar.radio(
    "Select Layout Mode",
    options=[
        "📌 Exact Input PDF Layout (Preserve 100% Original Coordinates & Positions)",
        "🏷️ Standardized HTML Flow Template (Render JSON in Responsive Cards)"
    ],
    index=0,
    help="Exact Input PDF Layout retains 100% pixel-perfect coordinates for every text block, heading, table, image, and banner from the original PDF."
)
use_template = (layout_mode_choice == "🏷️ Standardized HTML Flow Template (Render JSON in Responsive Cards)")

st.sidebar.markdown("---")
st.sidebar.markdown("### 🎨 Theme Customizer")

theme_preset = st.sidebar.selectbox(
    "Theme Color Preset",
    options=["Classic Navy (#1f497d)", "Emerald Green (#059669)", "Dark Charcoal (#1e293b)", "Crimson Red (#b91c1c)", "Royal Purple (#7c3aed)", "Custom Color"],
    index=0
)

color_map = {
    "Classic Navy (#1f497d)": "#1f497d",
    "Emerald Green (#059669)": "#059669",
    "Dark Charcoal (#1e293b)": "#1e293b",
    "Crimson Red (#b91c1c)": "#b91c1c",
    "Royal Purple (#7c3aed)": "#7c3aed"
}

if theme_preset == "Custom Color":
    primary_color = st.sidebar.color_picker("Primary Accent Color", value="#1f497d")
else:
    primary_color = color_map[theme_preset]

font_choice = st.sidebar.selectbox(
    "Font Family",
    options=[
        "Cambria, 'Times New Roman', serif",
        "Inter, -apple-system, sans-serif",
        "Georgia, serif",
        "Roboto, Arial, sans-serif",
        "Courier New, monospace"
    ],
    index=0
)

custom_title = st.sidebar.text_input("Document Title Override", value="")
custom_subtitle = st.sidebar.text_input("Header Subtitle", value="Universal Dynamic Document Report")

st.sidebar.markdown("---")
st.sidebar.markdown("### 👁️ Component Visibility (Flow Template)")
show_kv = st.sidebar.checkbox("Show Key-Value Pairs Grid", value=True)
show_tables = st.sidebar.checkbox("Show Data Tables", value=True)
show_sections = st.sidebar.checkbox("Show Content Section Boxes", value=True)
show_images = st.sidebar.checkbox("Show Extracted Images & Graphs", value=True)

st.sidebar.markdown("---")
st.sidebar.markdown("### 🖊️ Document-Type Adaptations")
show_footer_signatures = st.sidebar.checkbox(
    "Show Footer Signature Boxes", value=True,
    help="Turn off for document types (e.g. invoices, certificates) that don't need sign-off boxes."
)
show_badges = st.sidebar.checkbox(
    "Highlight Status Keywords in Tables", value=True,
    help="Colors table cells matching status words (danger/warning/success). Vocabulary is generic across document types, not lab-report-only."
)
banner_font_size_pt = st.sidebar.slider(
    "Banner Heading Min Font Size (pt) — Exact Layout Mode", min_value=9.0, max_value=20.0, value=12.5, step=0.5,
    help="Any white-on-color text at or above this size is treated as a banner heading. Tune per source PDF instead of assuming a fixed 13-14pt."
)

preview_height = st.sidebar.slider("Live HTML Preview Height (px)", min_value=500, max_value=1200, value=850, step=50)

# Build active theme config dictionary
theme_config = {
    "primary_color": primary_color,
    "table_header_bg": primary_color,
    "border_color": primary_color,
    "font_family": font_choice,
    "header_subtitle": custom_subtitle,
    "show_kv": show_kv,
    "show_tables": show_tables,
    "show_sections": show_sections,
    "show_images": show_images,
    "show_footer_signatures": show_footer_signatures,
    "show_badges": show_badges,
    "banner_font_size_pt": banner_font_size_pt
}
if custom_title.strip():
    theme_config["header_title"] = custom_title.strip()

# Header Banner
st.markdown("""
<div class="header-card">
    <div class="header-title">⚡ Universal Dynamic PDF & JSON Converter</div>
    <div class="header-subtitle">Upload ANY PDF format (lab report, invoice, certificate, bill, tax form, tech spec) to extract complete structured JSON data (key-values, tables, content boxes, images, graphs) and render custom-designed HTML, Word (.docx) & PDF documents.</div>
</div>
""", unsafe_allow_html=True)

# Session State Initialization
if "converted" not in st.session_state:
    st.session_state.converted = False
if "extracted_data" not in st.session_state:
    st.session_state.extracted_data = None
if "html_content" not in st.session_state:
    st.session_state.html_content = ""
if "file_name" not in st.session_state:
    st.session_state.file_name = ""
if "file_bytes" not in st.session_state:
    st.session_state.file_bytes = None

# Main Area - File Upload
st.subheader("📤 1. Upload Any Source PDF Document")
uploaded_file = st.file_uploader("Select PDF File", type=["pdf"], help="Upload any PDF document format.")

if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
    st.session_state.file_bytes = file_bytes
    file_size_kb = len(file_bytes) / 1024.0

    try:
        temp_doc = fitz.open(stream=file_bytes, filetype="pdf")
        page_count = len(temp_doc)
        temp_doc.close()
    except Exception:
        page_count = "Unknown"

    col_i1, col_i2, col_i3 = st.columns(3)
    with col_i1:
        st.info(f"**Filename:** `{uploaded_file.name}`")
    with col_i2:
        st.info(f"**File Size:** `{file_size_kb:.1f} KB`")
    with col_i3:
        st.info(f"**Total Pages:** `{page_count}`")

    st.markdown("---")
    st.subheader("⚡ 2. Extract Data to JSON & Render Dynamic HTML")

    if st.button("🚀 Extract Complete JSON & Render HTML Document"):
        with st.spinner("Processing PDF: extracting key-values, tables, images, graphs & rendering HTML..."):
            try:
                with tempfile.TemporaryDirectory() as tmp_dir:
                    tmp_dir_path = Path(tmp_dir)
                    pdf_input_path = tmp_dir_path / uploaded_file.name
                    with open(pdf_input_path, "wb") as f_in:
                        f_in.write(file_bytes)

                    # 1. Universal Extraction
                    extracted_data = extract_report_data(str(pdf_input_path))
                    
                    # 2. Render Dynamic HTML
                    doc_title_use = custom_title.strip() if custom_title.strip() else uploaded_file.name
                    if use_template:
                        html_content = generate_dynamic_template_html(extracted_data, doc_title=doc_title_use, theme_config=theme_config)
                    else:
                        doc_fitz = fitz.open(str(pdf_input_path))
                        html_content = render_exact_pdf_layout_html(doc_fitz, doc_title=doc_title_use, theme_config=theme_config)
                        doc_fitz.close()

                    st.session_state.converted = True
                    st.session_state.extracted_data = extracted_data
                    st.session_state.html_content = html_content
                    st.session_state.file_name = uploaded_file.name

                    st.success("🎉 Extraction & Dynamic HTML Rendering Completed Successfully!")
            except Exception as e:
                st.error(f"Error processing PDF document: {e}")

# Display Results & Interactive Inspectors when converted
if st.session_state.converted and st.session_state.extracted_data and st.session_state.file_bytes:
    extracted_data = st.session_state.extracted_data
    doc_summary = extracted_data.get("document_summary", {})

    # Live Theme & Layout Update Handler
    doc_title_use = custom_title.strip() if custom_title.strip() else st.session_state.file_name
    if use_template:
        html_content = generate_dynamic_template_html(extracted_data, doc_title=doc_title_use, theme_config=theme_config)
    else:
        doc_fitz = fitz.open(stream=st.session_state.file_bytes, filetype="pdf")
        html_content = render_exact_pdf_layout_html(doc_fitz, doc_title=doc_title_use, theme_config=theme_config)
        doc_fitz.close()

    st.session_state.html_content = html_content

    st.markdown("---")
    st.subheader("📊 Extracted Data & Real-Time HTML Preview")

    # Metrics Row
    col_m1, col_m2, col_m3, col_m4, col_m5 = st.columns(5)
    with col_m1:
        st.metric("Total Pages", doc_summary.get("total_pages", 0))
    with col_m2:
        st.metric("Key-Value Pairs", doc_summary.get("total_key_value_pairs", 0))
    with col_m3:
        st.metric("Data Tables", doc_summary.get("total_tables", 0))
    with col_m4:
        st.metric("Content Boxes", doc_summary.get("total_boxes", 0))
    with col_m5:
        st.metric("Images & Graphs", doc_summary.get("total_images_and_graphs", 0))

    # Main Tabs Inspector
    tab_preview, tab_kv, tab_tables, tab_images, tab_boxes, tab_json = st.tabs([
        "🌐 Live HTML Preview",
        "🔑 Key-Value Pairs",
        "📋 Data Tables",
        "🖼️ Images & Graphs",
        "📦 Content Boxes",
        "📄 Raw JSON Payload"
    ])

    # Tab 1: Live HTML Preview & Downloads
    with tab_preview:
        col_d1, col_d2, col_d3, col_d4 = st.columns(4)
        with col_d1:
            st.download_button(
                label="📥 Download HTML (`target.html`)",
                data=st.session_state.html_content.encode("utf-8"),
                file_name=f"{Path(st.session_state.file_name).stem}_target.html",
                mime="text/html"
            )
        with col_d2:
            st.download_button(
                label="📥 Download Extracted JSON (`data.json`)",
                data=json.dumps(extracted_data, indent=2, ensure_ascii=False).encode("utf-8"),
                file_name=f"{Path(st.session_state.file_name).stem}.json",
                mime="application/json"
            )
        with col_d3:
            try:
                docx_bytes = convert_json_to_docx(extracted_data, theme_config=theme_config)
                st.download_button(
                    label="📥 Download Word (`report.docx`)",
                    data=docx_bytes,
                    file_name=f"{Path(st.session_state.file_name).stem}_report.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e_docx:
                st.warning(f"Word export error: {e_docx}")
        with col_d4:
            st.info("💡 PDF output can be saved directly from browser print or CLI (`convert.py`)")

        st.markdown(f"**Rendered Dynamic HTML Preview (`Mode: {'Exact Original PDF Layout' if not use_template else 'Standardized Flow Template'}`, Theme Accent: `{primary_color}`):**")
        components.html(st.session_state.html_content, height=preview_height, scrolling=True)

    # Tab 2: Key-Value Pairs
    with tab_kv:
        kv_pairs = extracted_data.get("extracted_key_value_pairs", {})
        if kv_pairs:
            st.markdown(f"Found **{len(kv_pairs)}** extracted key-value header metadata pairs:")
            kv_table_data = [{"Label / Field Key": k, "Extracted Value": v} for k, v in kv_pairs.items()]
            st.dataframe(kv_table_data, use_container_width=True)
        else:
            st.info("No key-value header pairs detected in document text.")

    # Tab 3: Data Tables
    with tab_tables:
        tables_list = extracted_data.get("tables", [])
        if tables_list:
            st.markdown(f"Found **{len(tables_list)}** extracted table(s):")
            for t_idx, tab_item in enumerate(tables_list):
                headers = tab_item.get("headers", [])
                rows = tab_item.get("rows", [])
                page_n = tab_item.get("page", 1)
                st.markdown(f"#### Table {t_idx + 1} (Page {page_n})")
                if headers and rows:
                    import pandas as pd
                    try:
                        df = pd.DataFrame(rows, columns=headers if len(headers) == len(rows[0]) else None)
                        st.dataframe(df, use_container_width=True)
                    except Exception:
                        st.write("Headers:", headers)
                        st.write("Rows:", rows)
                else:
                    st.write("Raw Table Rows:", rows)
        else:
            st.info("No structured data tables detected in document.")

    # Tab 4: Images & Graphs
    with tab_images:
        images_list = extracted_data.get("images_and_graphs", [])
        if images_list:
            st.markdown(f"Found **{len(images_list)}** extracted raster images, logos, figures, and charts:")
            img_cols = st.columns(3)
            for idx, img in enumerate(images_list):
                col_target = img_cols[idx % 3]
                with col_target:
                    data_uri = img.get("data_uri")
                    page_n = img.get("page", 1)
                    img_type = img.get("type", "image").replace("_", " ").title()
                    w = img.get("width", 0)
                    h = img.get("height", 0)
                    if data_uri:
                        st.image(data_uri, caption=f"Page {page_n} • {img_type} ({w}×{h}px)")
        else:
            st.info("No embedded raster images or graphs found in document.")

    # Tab 5: Content Boxes
    with tab_boxes:
        boxes_list = extracted_data.get("content_sections", [])
        if boxes_list:
            st.markdown(f"Found **{len(boxes_list)}** extracted content boxes & text blocks:")
            for b_idx, box in enumerate(boxes_list):
                title = box.get("title", f"Box {b_idx + 1}")
                page_n = box.get("page", 1)
                content_text = box.get("content_text", [])
                with st.expander(f"📦 Page {page_n}: {title}"):
                    if isinstance(content_text, list):
                        for line in content_text:
                            st.write(line)
                    else:
                        st.write(content_text)
                    st.caption(f"Bounding Box: {box.get('bbox')}")
        else:
            st.info("No content section boxes detected.")

    # Tab 6: Raw JSON Viewer
    with tab_json:
        st.markdown("### 📄 Complete Extracted JSON Data")
        st.json(extracted_data)
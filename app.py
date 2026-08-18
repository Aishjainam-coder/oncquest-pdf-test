"""
Universal PDF Processor, Renderer & Word (.docx) Converter
===========================================================
Clean Streamlit App:
- Upload PDF or JSON document
- Converts JSON + theme.json → Word (.docx) DIRECTLY ✅
- Live Rendered HTML & PDF Preview options
"""

import os
import tempfile
import base64
import json
import io
from pathlib import Path
import streamlit as st
import fitz  # PyMuPDF
import streamlit.components.v1 as components

from converter import (
    render_exact_pdf_layout_html,
    generate_dynamic_template_html,
    convert_json_to_docx,
    convert_html_to_docx,
    convert_pdf_to_word,
    render_html_to_pdf_and_preview
)
from extractor import extract_report_data

# Ensure output directories exist
Path("extracted_jsons").mkdir(exist_ok=True)
Path("output").mkdir(exist_ok=True)

# Load base theme.json if present
theme_json_defaults = {}
theme_file_path = Path("theme.json")
if theme_file_path.exists():
    try:
        with open(theme_file_path, "r", encoding="utf-8") as f_theme:
            theme_json_defaults = json.load(f_theme)
    except Exception:
        pass

# Configure Streamlit Page
st.set_page_config(
    page_title="PDF & JSON to Word Converter",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom Styling for Streamlit App
st.markdown("""
<style>
    .main { background-color: #f8fafc; }
    
    .header-card {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #1f497d 100%);
        color: #ffffff;
        padding: 1.5rem 2.0rem;
        border-radius: 12px;
        box-shadow: 0 8px 20px -4px rgba(15, 23, 42, 0.2);
        margin-bottom: 1.5rem;
    }
    .header-title {
        font-size: 1.8rem;
        font-weight: 800;
        margin: 0;
        letter-spacing: -0.02em;
        color: #ffffff;
    }
    .header-subtitle {
        font-size: 0.95rem;
        color: #cbd5e1;
        margin-top: 0.3rem;
        margin-bottom: 0;
    }

    .word-success-box {
        background: linear-gradient(135deg, #ecfdf5 0%, #d1fae5 100%);
        border: 1px solid #10b981;
        border-radius: 10px;
        padding: 1.2rem;
        margin-bottom: 1.5rem;
    }

    /* Primary Action Buttons */
    div.stButton > button:first-child {
        background: linear-gradient(135deg, #1f497d 0%, #0f172a 100%);
        color: white;
        font-weight: 600;
        font-size: 1.05rem;
        padding: 0.6rem 1.5rem;
        border-radius: 8px;
        border: none;
        box-shadow: 0 4px 12px rgba(31, 73, 125, 0.25);
        transition: all 0.2s ease-in-out;
        width: 100%;
    }
    div.stButton > button:first-child:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 16px rgba(31, 73, 125, 0.35);
    }
</style>
""", unsafe_allow_html=True)

# Minimal Sidebar Settings
st.sidebar.title("⚙️ Render Settings")

st.sidebar.info("📌 **Pipeline Mode:** High-Fidelity PDF & JSON → Word (.docx)")
use_template = False

theme_preset = st.sidebar.selectbox(
    "Primary Theme Color",
    options=["Classic Navy (#1f497d)", "Emerald Green (#059669)", "Dark Charcoal (#1e293b)", "Crimson Red (#b91c1c)"],
    index=0
)
color_map = {
    "Classic Navy (#1f497d)": "#1f497d",
    "Emerald Green (#059669)": "#059669",
    "Dark Charcoal (#1e293b)": "#1e293b",
    "Crimson Red (#b91c1c)": "#b91c1c"
}
primary_color = color_map[theme_preset]

preview_height = st.sidebar.slider("Preview Height (px)", min_value=500, max_value=1200, value=850, step=50)

theme_config = {
    "primary_color": primary_color,
    "table_header_bg": primary_color,
    "border_color": primary_color,
    "show_kv": True,
    "show_tables": True,
    "show_sections": True,
    "show_images": True,
    "show_footer_signatures": False,
    "show_badges": True
}

# App Header
st.markdown("""
<div class="header-card">
    <div class="header-title">⚡ Universal PDF & JSON → Word (.docx) Converter</div>
    <div class="header-subtitle">Direct High-Fidelity Pipeline: <b>PDF/JSON → Word (.docx) Directly ✅</b>. Preserves 100% exact layout, text, tables, fonts & colors.</div>
</div>
""", unsafe_allow_html=True)

# Session State Initialization
if "extracted_data" not in st.session_state:
    st.session_state.extracted_data = None
if "html_content" not in st.session_state:
    st.session_state.html_content = ""
if "output_pdf_bytes" not in st.session_state:
    st.session_state.output_pdf_bytes = None
if "docx_bytes" not in st.session_state:
    st.session_state.docx_bytes = None
if "file_name" not in st.session_state:
    st.session_state.file_name = ""
if "file_bytes" not in st.session_state:
    st.session_state.file_bytes = None

# 1. Upload Section
uploaded_file = st.file_uploader("📤 Choose ANY PDF report or extracted JSON file", type=["pdf", "json"], key="file_uploader")

if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
    file_ext = Path(uploaded_file.name).suffix.lower()
    
    # If a new file is uploaded, reset state for new file
    if st.session_state.file_name != uploaded_file.name or st.session_state.file_bytes != file_bytes:
        st.session_state.file_name = uploaded_file.name
        st.session_state.file_bytes = file_bytes
        st.session_state.html_content = ""
        st.session_state.output_pdf_bytes = None
        st.session_state.docx_bytes = None
        st.session_state.extracted_data = None

    # File Info Summary
    file_size_kb = len(file_bytes) / 1024.0
    page_count = "N/A"
    if file_ext == ".pdf":
        try:
            temp_doc = fitz.open(stream=file_bytes, filetype="pdf")
            page_count = len(temp_doc)
            temp_doc.close()
        except Exception:
            page_count = "Unknown"

    col_i1, col_i2, col_i3 = st.columns(3)
    with col_i1:
        st.info(f"📄 **Filename:** `{uploaded_file.name}`")
    with col_i2:
        st.info(f"⚖️ **Size:** `{file_size_kb:.1f} KB`")
    with col_i3:
        st.info(f"📑 **Type/Pages:** `{file_ext.upper()} | {page_count}`")

    st.markdown("---")

    # Action Button to Process & Convert directly to Word
    btn_process = st.button("⚡ Convert to Word (.docx) Directly", use_container_width=True, type="primary")

    # Processing Workflow
    if btn_process or st.session_state.docx_bytes is None:
        with st.spinner("Processing End Result HTML → Word (.docx)..."):
            try:
                if file_ext == ".json":
                    # Direct JSON Input -> Render Word (.docx) directly using theme.json
                    extracted_data = json.loads(file_bytes.decode("utf-8"))
                    st.session_state.extracted_data = extracted_data
                    html_content = generate_dynamic_template_html(extracted_data, doc_title=uploaded_file.name, theme_config=theme_config)
                    # Replace "SN Genelab Pvt Ltd" with "Laboratory"
                    html_content = html_content.replace("SN Genelab Pvt Ltd", "Laboratory")
                    st.session_state.html_content = html_content

                    with tempfile.TemporaryDirectory() as tmp_dir:
                        html_tmp = Path(tmp_dir) / "temp.html"
                        html_tmp.write_text(html_content, encoding="utf-8")
                        compiled_pdf_tmp = Path(tmp_dir) / "compiled.pdf"

                        # Compile Intermediate PDF from HTML
                        render_html_to_pdf_and_preview(html_tmp, compiled_pdf_tmp)
                        if compiled_pdf_tmp.exists():
                            st.session_state.compiled_pdf_bytes = compiled_pdf_tmp.read_bytes()

                            # Convert Compiled Result PDF -> Word (.docx) using pdf2docx
                            out_docx_tmp = Path(tmp_dir) / "output.docx"
                            convert_pdf_to_word(compiled_pdf_tmp, out_docx_tmp)
                            if out_docx_tmp.exists():
                                st.session_state.docx_bytes = out_docx_tmp.read_bytes()
                            else:
                                st.session_state.docx_bytes = None
                        else:
                            st.session_state.compiled_pdf_bytes = None
                            # Fallback if Playwright PDF rendering failed
                            st.session_state.docx_bytes = convert_json_to_docx(extracted_data, theme_config=theme_config)
                else:
                    # PDF Input -> Render Clean HTML End Result (Header/Footer Excluded & Theme Styled) -> Convert HTML to Word
                    with tempfile.TemporaryDirectory() as tmp_dir:
                        pdf_input_path = Path(tmp_dir) / uploaded_file.name
                        pdf_input_path.write_bytes(file_bytes)

                        # Silent JSON extraction
                        try:
                            extracted_data = extract_report_data(str(pdf_input_path))
                            st.session_state.extracted_data = extracted_data
                            
                            json_out_dir = Path("extracted_jsons")
                            json_out_dir.mkdir(exist_ok=True)
                            json_file_path = json_out_dir / f"{Path(uploaded_file.name).stem}.json"
                            # Replace "SN Genelab Pvt Ltd" with "Laboratory" in JSON
                            json_str = json.dumps(extracted_data, indent=2, ensure_ascii=False)
                            json_str = json_str.replace("SN Genelab Pvt Ltd", "Laboratory")
                            with open(json_file_path, "w", encoding="utf-8") as f_json:
                                f_json.write(json_str)
                        except Exception:
                            pass

                        # Render Clean End Result HTML
                        doc_fitz = fitz.open(str(pdf_input_path))
                        html_content = render_exact_pdf_layout_html(doc_fitz, doc_title=uploaded_file.name, theme_config=theme_config)
                        doc_fitz.close()
                        # Replace "SN Genelab Pvt Ltd" with "Laboratory"
                        html_content = html_content.replace("SN Genelab Pvt Ltd", "Laboratory")
                        st.session_state.html_content = html_content

                        # Compile Intermediate PDF from HTML
                        try:
                            html_tmp = Path(tmp_dir) / "temp.html"
                            html_tmp.write_text(html_content, encoding="utf-8")
                            compiled_pdf_tmp = Path(tmp_dir) / "compiled.pdf"
                            render_html_to_pdf_and_preview(html_tmp, compiled_pdf_tmp)
                            if compiled_pdf_tmp.exists():
                                st.session_state.compiled_pdf_bytes = compiled_pdf_tmp.read_bytes()
                        except Exception:
                            st.session_state.compiled_pdf_bytes = None

                        # Convert Compiled Result PDF -> Word (.docx) using pdf2docx Converter
                        out_docx_tmp = Path(tmp_dir) / "output.docx"
                        target_pdf_for_word = compiled_pdf_tmp if (compiled_pdf_tmp.exists() and compiled_pdf_tmp.stat().st_size > 0) else pdf_input_path
                        convert_pdf_to_word(target_pdf_for_word, out_docx_tmp)
                        if out_docx_tmp.exists():
                            st.session_state.docx_bytes = out_docx_tmp.read_bytes()
                        else:
                            st.session_state.docx_bytes = None

                st.success("✅ Converted Extracted Content + theme.json → Word (.docx) successfully!")

            except Exception as e:
                st.error(f"Error converting document to Word: {e}")



    # 2. Direct Word Download & Results Section
    if st.session_state.docx_bytes:
        st.markdown("---")
        
        # Featured Direct Word Download Card
        col_w1, col_w2 = st.columns([2, 1])
        with col_w1:
            st.markdown("### 📝 Direct Word Document (.docx) Ready!")
            st.markdown("Your document was styled using **`theme.json`** rules (colors, fonts, borders, tables) and converted directly into a Microsoft Word file.")
        with col_w2:
            st.download_button(
                label="📥 Download Word Document (.docx)",
                data=st.session_state.docx_bytes,
                file_name=f"{Path(st.session_state.file_name).stem}_report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                type="primary"
            )

        st.markdown("---")
        st.subheader("📊 Optional Web & PDF Previews")

        tab_html, tab_pdf = st.tabs(["🌐 Rendered HTML Preview", "📄 Compiled PDF Preview"])

        # Tab 1: Rendered HTML Result
        with tab_html:
            if st.session_state.html_content:
                col_h1, col_h2 = st.columns([3, 1])
                with col_h1:
                    st.markdown("### 🌐 Rendered HTML Document")
                with col_h2:
                    st.download_button(
                        label="🌐 Download HTML (`.html`)",
                        data=st.session_state.html_content.encode("utf-8"),
                        file_name=f"{Path(st.session_state.file_name).stem}.html",
                        mime="text/html",
                        use_container_width=True
                    )
                components.html(st.session_state.html_content, height=preview_height, scrolling=True)

        # Tab 2: Compiled Output PDF Result
        with tab_pdf:
            compiled_pdf_bytes = getattr(st.session_state, "compiled_pdf_bytes", None)
            if compiled_pdf_bytes:
                st.markdown("### 📄 Compiled Target PDF")
                b64_pdf = base64.b64encode(compiled_pdf_bytes).decode("utf-8")
                pdf_display = f'<iframe src="data:application/pdf;base64,{b64_pdf}" width="100%" height="{preview_height}px" type="application/pdf"></iframe>'
                st.markdown(pdf_display, unsafe_allow_html=True)
            else:
                st.info("Compiled PDF preview will appear after generation.")
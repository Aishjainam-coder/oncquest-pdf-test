"""
Universal PDF Processor, Renderer & Word (.docx) Converter
===========================================================
Clean Streamlit App:
- Upload PDF
- Extracts JSON in background (saved to extracted_jsons/ folder)
- Live Rendered HTML Result & Output PDF Result
- 1-Click Convert Result PDF to Word (.docx)
"""

import os
import tempfile
import base64
import json
from pathlib import Path
import streamlit as st
import fitz  # PyMuPDF
import streamlit.components.v1 as components

from converter import (
    render_exact_pdf_layout_html,
    generate_dynamic_template_html,
    convert_json_to_docx,
    convert_html_to_docx,
    render_html_to_pdf_and_preview
)
from extractor import extract_report_data

# Ensure output directories exist
Path("extracted_jsons").mkdir(exist_ok=True)
Path("output").mkdir(exist_ok=True)

# Configure Streamlit Page
st.set_page_config(
    page_title="PDF Processor, HTML & Word Converter",
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

st.sidebar.info("📌 **PDF Layout Mode:** Exact Input PDF Layout (Preserve Coordinates)")
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
    "show_footer_signatures": True,
    "show_badges": True
}

# App Header
st.markdown("""
<div class="header-card">
    <div class="header-title">⚡ Universal PDF Processor & Document Renderer</div>
    <div class="header-subtitle">Upload any PDF document to automatically process structured data, view Rendered HTML & Output PDF results, and convert the result to Word (.docx).</div>
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
uploaded_file = st.file_uploader("📤 Choose ANY PDF report, lab test, invoice, or document", type=["pdf"], key="pdf_uploader")

if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
    
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
        st.info(f"📑 **Pages:** `{page_count}`")

    st.markdown("---")

    # Action Button to Process PDF
    btn_process = st.button("⚡ Process PDF & Render Results", use_container_width=True, type="primary")

    # Processing Workflow
    if btn_process or st.session_state.html_content == "":
        with st.spinner("Extracting JSON silently & rendering HTML and Output PDF..."):
            try:
                with tempfile.TemporaryDirectory() as tmp_dir:
                    pdf_input_path = Path(tmp_dir) / uploaded_file.name
                    with open(pdf_input_path, "wb") as f_in:
                        f_in.write(file_bytes)

                    # Step 1: Extract structured JSON data
                    extracted_data = extract_report_data(str(pdf_input_path))
                    st.session_state.extracted_data = extracted_data

                    # Save extracted JSON silently in background folder
                    try:
                        json_out_dir = Path("extracted_jsons")
                        json_out_dir.mkdir(exist_ok=True)
                        json_file_path = json_out_dir / f"{Path(uploaded_file.name).stem}.json"
                        with open(json_file_path, "w", encoding="utf-8") as f_json:
                            json.dump(extracted_data, f_json, indent=2, ensure_ascii=False)
                    except Exception as e_json:
                        st.warning(f"Could not save JSON to extracted_jsons folder: {e_json}")

                    # Step 2: Render HTML Content
                    doc_title_use = uploaded_file.name
                    if use_template:
                        html_content = generate_dynamic_template_html(extracted_data, doc_title=doc_title_use, theme_config=theme_config)
                    else:
                        doc_fitz = fitz.open(str(pdf_input_path))
                        html_content = render_exact_pdf_layout_html(doc_fitz, doc_title=doc_title_use, theme_config=theme_config)
                        doc_fitz.close()

                    st.session_state.html_content = html_content

                    # Step 3: Render Output PDF from HTML
                    tmp_html_file = Path(tmp_dir) / f"{Path(uploaded_file.name).stem}_render.html"
                    tmp_pdf_file = Path(tmp_dir) / f"{Path(uploaded_file.name).stem}_result.pdf"
                    tmp_html_file.write_text(html_content, encoding="utf-8")

                    render_html_to_pdf_and_preview(tmp_html_file, tmp_pdf_file)

                    if tmp_pdf_file.exists():
                        with open(tmp_pdf_file, "rb") as f_pdf:
                            st.session_state.output_pdf_bytes = f_pdf.read()
                    else:
                        st.session_state.output_pdf_bytes = file_bytes

                    st.success("✅ PDF processed successfully! JSON saved to `extracted_jsons/` folder.")

            except Exception as e:
                st.error(f"Error processing PDF document: {e}")

    # 2. Result Section (Web UI Output)
    if st.session_state.html_content:
        st.markdown("---")
        st.subheader("📊 Conversion Results & Output Preview")

        tab_html, tab_pdf = st.tabs(["🌐 Rendered HTML Result", "📄 Output PDF Result"])

        # Tab 1: Rendered HTML Result
        with tab_html:
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

        # Tab 2: Output PDF Result & Word Conversion
        with tab_pdf:
            col_p1, col_p2 = st.columns([1, 1])
            with col_p1:
                st.markdown("### 📄 Output PDF Document")
            with col_p2:
                if st.session_state.output_pdf_bytes:
                    st.download_button(
                        label="📄 Download Output PDF (`.pdf`)",
                        data=st.session_state.output_pdf_bytes,
                        file_name=f"{Path(st.session_state.file_name).stem}_output.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )

            st.markdown("---")

            # Word (.docx) Conversion Box under Result PDF
            col_w1, col_w2 = st.columns([1, 1])
            with col_w1:
                st.markdown("#### 📝 Convert Result PDF to Word (.docx)")
                btn_conv_word = st.button("📝 Convert Result PDF to Word (.docx)", use_container_width=True, type="primary")

                if btn_conv_word:
                    with st.spinner("Converting Document Data to Word (.docx)..."):
                        try:
                            if st.session_state.extracted_data:
                                docx_data = convert_json_to_docx(st.session_state.extracted_data, theme_config=theme_config)
                            else:
                                docx_data = convert_html_to_docx(st.session_state.html_content, theme_config=theme_config)
                            st.session_state.docx_bytes = docx_data
                            st.success("✅ Converted Document to Word (.docx) successfully with perfect table alignment!")
                        except Exception as err_w:
                            st.error(f"Error converting to Word: {err_w}")

            with col_w2:
                if st.session_state.docx_bytes:
                    st.markdown("#### 📥 Ready for Download")
                    st.download_button(
                        label="📥 Download Word Document (`.docx`)",
                        data=st.session_state.docx_bytes,
                        file_name=f"{Path(st.session_state.file_name).stem}_result.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )

            st.markdown("---")

            # Embedded Output PDF Display
            if st.session_state.output_pdf_bytes:
                b64_pdf = base64.b64encode(st.session_state.output_pdf_bytes).decode("utf-8")
                pdf_display = f'<iframe src="data:application/pdf;base64,{b64_pdf}" width="100%" height="{preview_height}px" type="application/pdf"></iframe>'
                st.markdown(pdf_display, unsafe_allow_html=True)
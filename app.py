"""
Universal PDF Processor, Renderer & Word (.docx) Converter
===========================================================
Clean Streamlit App:
- Upload PDF or JSON document
- Converts JSON + theme.json → Word (.docx) DIRECTLY ✅
- Live Rendered HTML & PDF Preview options
"""

import pymupdf
pymupdf._g_out_message = None

import os
import tempfile
import base64
import json
import io
import time
from pathlib import Path
import streamlit as st
import pymupdf as fitz  # PyMuPDF

from converter import (
    render_exact_pdf_layout_html,
    generate_dynamic_template_html,
    convert_json_to_docx,
    convert_html_to_docx,
    convert_pdf_to_word,
    convert_pdf_via_pdf2docx,
    render_html_to_pdf_and_preview
)
from extractor import extract_report_data

BASE_DIR = Path(__file__).resolve().parent

def format_time_duration(seconds: float) -> str:
    """Format seconds into readable minutes and seconds (e.g., '16m 16.14s' or '12.45s')."""
    if seconds is None:
        return "0.00s"
    try:
        seconds = float(seconds)
    except Exception:
        return str(seconds)
    total_secs = int(seconds)
    hours = total_secs // 3600
    mins = (total_secs % 3600) // 60
    rem_secs = seconds % 60
    if hours > 0:
        return f"{hours}h {mins}m {rem_secs:.2f}s"
    if mins > 0:
        return f"{mins}m {rem_secs:.2f}s"
    return f"{rem_secs:.2f}s"

# Ensure output directories exist
(BASE_DIR / "extracted_jsons").mkdir(exist_ok=True)
(BASE_DIR / "output").mkdir(exist_ok=True)

# Load base theme.json if present
theme_json_defaults = {}
theme_file_path = BASE_DIR / "theme.json"
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
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    }
    
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
        font-family: 'Inter', sans-serif;
    }
    .header-subtitle {
        font-size: 0.95rem;
        color: #cbd5e1;
        margin-top: 0.3rem;
        margin-bottom: 0;
        font-family: 'Inter', sans-serif;
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
        font-family: 'Inter', sans-serif;
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
if "compiled_pdf_bytes" not in st.session_state:
    st.session_state.compiled_pdf_bytes = None
if "docx_bytes" not in st.session_state:
    st.session_state.docx_bytes = None
if "file_name" not in st.session_state:
    st.session_state.file_name = ""
if "file_bytes" not in st.session_state:
    st.session_state.file_bytes = None
if "time_taken" not in st.session_state:
    st.session_state.time_taken = None
if "step_times" not in st.session_state:
    st.session_state.step_times = {}

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
        st.session_state.compiled_pdf_bytes = None
        st.session_state.docx_bytes = None
        st.session_state.extracted_data = None
        st.session_state.time_taken = None
        st.session_state.step_times = {}

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

    if st.session_state.time_taken is not None:
        col_i1, col_i2, col_i3, col_i4 = st.columns(4)
        with col_i1:
            st.info(f"📄 **Filename:** `{uploaded_file.name}`")
        with col_i2:
            st.info(f"⚖️ **Size:** `{file_size_kb:.1f} KB`")
        with col_i3:
            st.info(f"📑 **Type/Pages:** `{file_ext.upper()} | {page_count}`")
        with col_i4:
            st.success(f"⏱️ **Time Taken:** `{format_time_duration(st.session_state.time_taken)}`")

        # Step Execution Times Breakdown
        if st.session_state.step_times:
            st.markdown("##### ⏱️ Step-by-Step Processing Times")
            step_cols = st.columns(len(st.session_state.step_times))
            for col, (s_name, s_time) in zip(step_cols, st.session_state.step_times.items()):
                with col:
                    st.info(f"**{s_name}**\n\n⏱️ `{format_time_duration(s_time)}`")
    else:
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
        start_time = time.perf_counter()
        step_times = {}
        print(f"\n{'='*60}", flush=True)
        print(f"[*] Starting Document Processing: {uploaded_file.name}", flush=True)
        print(f"{'='*60}", flush=True)

        with st.status("⚡ Converting document to Word (.docx)...", expanded=True) as status_box:
            try:
                if file_ext == ".json":
                    # Direct JSON Input -> Render Word (.docx) directly using theme.json
                    t0 = time.perf_counter()
                    step1_ph = st.empty()
                    step1_ph.write("🔍 **Step 1/3:** Loading and parsing JSON structure... ⏳ *(processing)*")
                    print(f"[*] [Step 1/3] Loading JSON structure for {uploaded_file.name}...", flush=True)
                    extracted_data = json.loads(file_bytes.decode("utf-8"))
                    st.session_state.extracted_data = extracted_data
                    dur_step1 = time.perf_counter() - t0
                    step_times["🔍 Step 1: JSON Parsing"] = dur_step1
                    step1_ph.write(f"🔍 **Step 1/3:** Loading and parsing JSON structure — ⏱️ **{format_time_duration(dur_step1)}**")

                    t1 = time.perf_counter()
                    step2_ph = st.empty()
                    step2_ph.write("🎨 **Step 2/3:** Generating themed HTML layout from JSON... ⏳ *(processing)*")
                    print(f"[*] [Step 2/3] Generating themed HTML template...", flush=True)
                    html_content = generate_dynamic_template_html(extracted_data, doc_title=uploaded_file.name, theme_config=theme_config)
                    html_content = html_content.replace("SN Genelab Pvt Ltd", "Laboratory")
                    st.session_state.html_content = html_content
                    dur_step2 = time.perf_counter() - t1
                    step_times["🎨 Step 2: HTML Layout"] = dur_step2
                    step2_ph.write(f"🎨 **Step 2/3:** Generating themed HTML layout from JSON — ⏱️ **{format_time_duration(dur_step2)}**")

                    with tempfile.TemporaryDirectory() as tmp_dir:
                        html_tmp = Path(tmp_dir) / "temp.html"
                        html_tmp.write_text(html_content, encoding="utf-8")
                        compiled_pdf_tmp = Path(tmp_dir) / "compiled.pdf"

                        t2 = time.perf_counter()
                        step3_ph = st.empty()
                        step3_ph.write("🌐 **Step 3/3:** Compiling HTML to PDF and converting to Word (.docx)... ⏳ *(processing)*")
                        print(f"[*] [Step 3/3] Compiling intermediate PDF via Playwright...", flush=True)
                        render_html_to_pdf_and_preview(html_tmp, compiled_pdf_tmp)

                        if compiled_pdf_tmp.exists():
                            st.session_state.compiled_pdf_bytes = compiled_pdf_tmp.read_bytes()
                            docx_tmp = Path(tmp_dir) / "output.docx"
                            print(f"[*] Converting compiled PDF to Word (.docx) via pdf2docx...", flush=True)
                            convert_pdf_via_pdf2docx(str(compiled_pdf_tmp), str(docx_tmp))
                            dur_step3 = time.perf_counter() - t2
                            step_times["📝 Step 3: DOCX Generation"] = dur_step3
                            step3_ph.write(f"🌐 **Step 3/3:** Compiling HTML to PDF and converting to Word (.docx) — ⏱️ **{format_time_duration(dur_step3)}**")
                            if docx_tmp.exists():
                                st.session_state.docx_bytes = docx_tmp.read_bytes()
                                print(f"[+] DOCX generation successful ({len(st.session_state.docx_bytes)} bytes)!", flush=True)
                            else:
                                st.session_state.docx_bytes = None
                        else:
                            st.session_state.compiled_pdf_bytes = None
                            st.session_state.docx_bytes = None
                else:
                    # PDF Input -> Render Clean HTML End Result -> Convert HTML to Word
                    with tempfile.TemporaryDirectory() as tmp_dir:
                        pdf_input_path = Path(tmp_dir) / uploaded_file.name
                        pdf_input_path.write_bytes(file_bytes)

                        # Step 1: Extract structured JSON from PDF
                        t0 = time.perf_counter()
                        step1_ph = st.empty()
                        step1_ph.write("🔍 **Step 1/4:** Extracting text, tables, and styles from PDF... ⏳ *(processing)*")
                        print(f"[*] [Step 1/4] Extracting text, tables, and styles from PDF...", flush=True)
                        try:
                            extracted_data = extract_report_data(str(pdf_input_path), auto_save_docx=False)
                            st.session_state.extracted_data = extracted_data
                            
                            json_out_dir = BASE_DIR / "extracted_jsons"
                            json_out_dir.mkdir(exist_ok=True)
                            json_file_path = json_out_dir / f"{Path(uploaded_file.name).stem}.json"
                            json_str = json.dumps(extracted_data, indent=2, ensure_ascii=False)
                            json_str = json_str.replace("SN Genelab Pvt Ltd", "Laboratory")
                            with open(json_file_path, "w", encoding="utf-8") as f_json:
                                f_json.write(json_str)
                            print(f"   [+] Extracted JSON saved to: {json_file_path}", flush=True)
                        except Exception as e_ext:
                            print(f"   [!] Note on JSON extraction: {e_ext}", flush=True)
                        dur_step1 = time.perf_counter() - t0
                        step_times["🔍 Step 1: Extract"] = dur_step1
                        step1_ph.write(f"🔍 **Step 1/4:** Extracting text, tables, and styles from PDF — ⏱️ **{format_time_duration(dur_step1)}**")

                        # Step 2: Render Clean End Result HTML
                        t1 = time.perf_counter()
                        step2_ph = st.empty()
                        step2_ph.write("🎨 **Step 2/4:** Rendering styled HTML document layout... ⏳ *(processing)*")
                        print(f"[*] [Step 2/4] Rendering styled HTML document layout...", flush=True)
                        with fitz.open(str(pdf_input_path)) as doc_fitz:
                            html_content = render_exact_pdf_layout_html(doc_fitz, doc_title=uploaded_file.name, theme_config=theme_config)
                        html_content = html_content.replace("SN Genelab Pvt Ltd", "Laboratory")
                        st.session_state.html_content = html_content
                        dur_step2 = time.perf_counter() - t1
                        step_times["🎨 Step 2: HTML Layout"] = dur_step2
                        step2_ph.write(f"🎨 **Step 2/4:** Rendering styled HTML document layout — ⏱️ **{format_time_duration(dur_step2)}**")

                        # Step 3: Compile HTML to PDF via Playwright
                        t2 = time.perf_counter()
                        step3_ph = st.empty()
                        step3_ph.write("🌐 **Step 3/4:** Compiling HTML to PDF preview via Chromium... ⏳ *(processing)*")
                        print(f"[*] [Step 3/4] Compiling HTML to PDF preview via Chromium...", flush=True)
                        html_tmp = Path(tmp_dir) / "temp.html"
                        html_tmp.write_text(html_content, encoding="utf-8")
                        compiled_pdf_tmp = Path(tmp_dir) / "compiled.pdf"
                        render_html_to_pdf_and_preview(html_tmp, compiled_pdf_tmp)
                        dur_step3 = time.perf_counter() - t2
                        step_times["🌐 Step 3: PDF Compile"] = dur_step3
                        step3_ph.write(f"🌐 **Step 3/4:** Compiling HTML to PDF preview via Chromium — ⏱️ **{format_time_duration(dur_step3)}**")

                        # Step 4: Convert compiled PDF to Word (.docx)
                        if compiled_pdf_tmp.exists():
                            st.session_state.compiled_pdf_bytes = compiled_pdf_tmp.read_bytes()
                            t3 = time.perf_counter()
                            step4_ph = st.empty()
                            step4_ph.write("📝 **Step 4/4:** Reconstructing Word (.docx) with exact layout & styling... ⏳ *(processing)*")
                            print(f"[*] [Step 4/4] Converting compiled PDF to Word (.docx) via pdf2docx...", flush=True)
                            docx_tmp = Path(tmp_dir) / "output.docx"
                            convert_pdf_via_pdf2docx(str(compiled_pdf_tmp), str(docx_tmp))
                            dur_step4 = time.perf_counter() - t3
                            step_times["📝 Step 4: DOCX Conversion"] = dur_step4
                            step4_ph.write(f"📝 **Step 4/4:** Reconstructing Word (.docx) with exact layout & styling — ⏱️ **{format_time_duration(dur_step4)}**")
                            if docx_tmp.exists():
                                st.session_state.docx_bytes = docx_tmp.read_bytes()
                                print(f"[+] DOCX generation successful ({len(st.session_state.docx_bytes)} bytes)!", flush=True)
                            else:
                                st.session_state.docx_bytes = None
                        else:
                            st.session_state.compiled_pdf_bytes = None
                            st.session_state.docx_bytes = None

                elapsed_total = time.perf_counter() - start_time
                st.session_state.time_taken = elapsed_total
                st.session_state.step_times = step_times

                formatted_total = format_time_duration(elapsed_total)
                status_box.update(label=f"✅ Conversion Completed Successfully in {formatted_total}!", state="complete", expanded=False)
                print(f"[+] Pipeline Completed Successfully for {uploaded_file.name} in {formatted_total} ({elapsed_total:.2f}s)!\n", flush=True)
                
                breakdown_str = " | ".join([f"{k}: **{format_time_duration(v)}**" for k, v in step_times.items()])
                st.success(f"⏱️ **Total Time Taken:** `{formatted_total}` ({breakdown_str})")
                st.rerun()

            except Exception as e:
                status_box.update(label=f"❌ Error during conversion: {e}", state="error", expanded=True)
                print(f"[!] Error converting document: {e}", flush=True)
                st.error(f"Error converting document to Word: {e}")

    # 2. Direct Word Download & Results Section
    if st.session_state.docx_bytes:
        st.markdown("---")
        
        # Featured Direct Word Download Card
        col_w1, col_w2 = st.columns([2, 1])
        with col_w1:
            st.markdown("### 📝 Direct Word Document (.docx) Ready!")
            time_display = f"⏱️ **Generated in `{format_time_duration(st.session_state.time_taken)}`** &nbsp;|&nbsp; " if st.session_state.time_taken else ""
            st.markdown(f"{time_display}Your document was styled using **`theme.json`** rules (colors, fonts, borders, tables) and converted directly into a Microsoft Word file.")
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
                st.iframe(st.session_state.html_content, height=preview_height)

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
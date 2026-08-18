import fitz
import sys
import re
import shutil
import os
from pathlib import Path

# Ensure UTF-8 output
sys.stdout.reconfigure(encoding='utf-8')

workspace_dir = Path(r"c:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main")
downloads_dir = Path(r"c:\Users\aishwarya.jain\Downloads")
backup_dir = workspace_dir / "backup_pdfs"
backup_dir.mkdir(exist_ok=True)

# Find all PDFs
pdf_paths = (
    list(workspace_dir.glob("output/**/*.pdf")) +
    list(workspace_dir.glob("outsourcing pdf/**/*.pdf")) +
    list(downloads_dir.glob("*.pdf"))
)

# Deduplicate
unique_pdfs = []
seen = set()
for p in pdf_paths:
    res = p.resolve()
    if res not in seen:
        seen.add(res)
        unique_pdfs.append(p)

print(f"[+] Found {len(unique_pdfs)} unique PDF files to scan.")

for pdf_path in unique_pdfs:
    if not pdf_path.exists():
        continue
        
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        print(f"[*] Error opening {pdf_path.name}: {e}")
        continue
        
    page = doc[0]
    blocks = page.get_text("dict")["blocks"]
    
    # Classify format
    format_type = None
    rects_to_redact = []
    text_insertions = []  # list of tuples (rect, text, font_path, font_size, color_rgb, align)
    
    # 1. Format 1: Dulal Naha / exact_position_test (centered 2-line test name)
    span_title = None
    span_subtitle = None
    for b in blocks:
        if "lines" not in b:
            continue
        for l in b["lines"]:
            for s in l["spans"]:
                if "Liquidseq Actionable Genomic Profiling Panel" in s["text"]:
                    span_title = s
                elif "On Illumina Novaseq 6000 Platform" in s["text"]:
                    span_subtitle = s
                    
    if span_title and span_subtitle:
        format_type = 1
        rect1 = fitz.Rect(span_title["bbox"])
        rect2 = fitz.Rect(span_subtitle["bbox"])
        rects_to_redact = [rect1, rect2]
        
        color_int = span_title["color"]
        r = ((color_int >> 16) & 255) / 255.0
        g = ((color_int >> 8) & 255) / 255.0
        b = (color_int & 255) / 255.0
        color_rgb = (r, g, b)
        
        rect1_expanded = fitz.Rect(rect1)
        rect1_expanded.y0 -= 3
        rect1_expanded.y1 += 3
        
        text_insertions = [(
            rect1_expanded,
            "TEST NAME",
            "C:/Windows/Fonts/cambriaz.ttf",  # Cambria Bold Italic
            span_title["size"],
            color_rgb,
            1  # center aligned
        )]

    # 2. Format 2: Birendra / Basanti (white test name on dark background banner)
    if not format_type:
        target_span = None
        for b in blocks:
            if "lines" not in b:
                continue
            for l in b["lines"]:
                for s in l["spans"]:
                    y0 = s["bbox"][1]
                    if 170 <= y0 <= 185:
                        if s["text"].strip() == "PERFORMED." or "HEREDITARY" in s["text"]:
                            target_span = s
                            break
                if target_span:
                    break
        if target_span:
            format_type = 2
            rect = fitz.Rect(target_span["bbox"])
            rects_to_redact = [rect]
            
            color_int = target_span["color"]
            r = ((color_int >> 16) & 255) / 255.0
            g = ((color_int >> 8) & 255) / 255.0
            b = (color_int & 255) / 255.0
            color_rgb = (r, g, b)
            
            rect_expanded = fitz.Rect(rect)
            rect_expanded.y0 -= 3
            rect_expanded.y1 += 3
            
            text_insertions = [(
                rect_expanded,
                "TEST NAME",
                "C:/Windows/Fonts/cambriab.ttf",  # Cambria Bold
                target_span["size"],
                color_rgb,
                1  # center aligned
            )]

    # 3. Format 3: vendor single line patient metadata containing ": Test Name "
    if not format_type:
        target_span = None
        for b in blocks:
            if "lines" not in b:
                continue
            for l in b["lines"]:
                for s in l["spans"]:
                    if ": Test Name " in s["text"] and len(s["text"]) > 25:
                        target_span = s
                        break
                if target_span:
                    break
        if target_span:
            format_type = 3
            rect = fitz.Rect(target_span["bbox"])
            rects_to_redact = [rect]
            
            orig_text = target_span["text"]
            new_text = re.sub(r"(: Test Name\s*)(.*)", r"\1TEST NAME", orig_text)
            
            color_int = target_span["color"]
            r = ((color_int >> 16) & 255) / 255.0
            g = ((color_int >> 8) & 255) / 255.0
            b = (color_int & 255) / 255.0
            color_rgb = (r, g, b)
            
            rect_expanded = fitz.Rect(rect)
            rect_expanded.y0 -= 3
            rect_expanded.y1 += 3
            
            text_insertions = [(
                rect_expanded,
                new_text,
                "C:/Windows/Fonts/cambria.ttc",  # Cambria Regular
                target_span["size"],
                color_rgb,
                0  # left aligned
            )]

    # 4. Format 4: vendor separate spans (Test Name label and value)
    if not format_type:
        label_span = None
        value_span = None
        for b in blocks:
            if "lines" not in b:
                continue
            for l in b["lines"]:
                for s in l["spans"]:
                    if s["text"].strip() == "Test Name" and abs(s["size"] - 4.56) < 0.1:
                        label_span = s
                        break
                if label_span:
                    break
        if label_span:
            y_label = label_span["bbox"][1]
            for b in blocks:
                if "lines" not in b:
                    continue
                for l in b["lines"]:
                    for s in l["spans"]:
                        y_s = s["bbox"][1]
                        if abs(y_s - y_label) < 1.0 and "ORION" in s["text"]:
                            value_span = s
                            break
                    if value_span:
                        break
            if value_span:
                format_type = 4
                rect = fitz.Rect(value_span["bbox"])
                rects_to_redact = [rect]
                
                color_int = value_span["color"]
                r = ((color_int >> 16) & 255) / 255.0
                g = ((color_int >> 8) & 255) / 255.0
                b = (color_int & 255) / 255.0
                color_rgb = (r, g, b)
                
                rect_expanded = fitz.Rect(rect)
                rect_expanded.y0 -= 3
                rect_expanded.y1 += 3
                
                text_insertions = [(
                    rect_expanded,
                    "TEST NAME",
                    "C:/Windows/Fonts/cambria.ttc",  # Cambria Regular
                    value_span["size"],
                    color_rgb,
                    0  # left aligned
                )]

    # Apply modifications
    if format_type:
        print(f"[+] Processing {pdf_path.name} (Format {format_type})...")
        
        # Backup
        backup_file = backup_dir / pdf_path.name
        if not backup_file.exists():
            shutil.copy2(pdf_path, backup_file)
            print(f"  [Backup] Copied to {backup_file}")
            
        # Redact old text
        for r in rects_to_redact:
            page.add_redact_annot(r)
        page.apply_redactions(graphics=0)
        
        # Insert replacement
        for rect, text, font_path, size, color_rgb, align in text_insertions:
            ret = page.insert_textbox(
                rect,
                text,
                fontname="custom-font",
                fontfile=font_path,
                fontsize=size,
                color=color_rgb,
                align=align
            )
            if ret < 0:
                print(f"  [Warning] Textbox insertion returned overflow: {ret}")
            else:
                print(f"  [Text] Successfully inserted: '{text}'")
                
        # Save to a temporary file first
        tmp_path = pdf_path.with_suffix(".tmp")
        try:
            doc.save(tmp_path)
            doc.close()
            # Replace original file with the modified one
            if tmp_path.exists():
                shutil.move(str(tmp_path), str(pdf_path))
                print(f"  [Save] Successfully saved changes in place.")
        except Exception as e:
            print(f"  [Error] Failed to save {pdf_path.name}: {e}")
            if tmp_path.exists():
                try:
                    os.remove(tmp_path)
                except:
                    pass
    else:
        print(f"[-] Skipped {pdf_path.name} (no test-name field found).")
        doc.close()

print("\n[+] All PDF replacements completed successfully!")

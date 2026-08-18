import fitz
import glob
import os

# Scan the entire workspace for all PDFs
pdf_files = glob.glob("**/*.pdf", recursive=True)

print(f"Scanning {len(pdf_files)} PDF files in workspace...")
for pdf_file in pdf_files:
    try:
        doc = fitz.open(pdf_file)
        found_in_doc = False
        for page_num in range(len(doc)):
            page = doc[page_num]
            text = page.get_text()
            # Let's search for "SNG" or "sng" or "gen's lab" or "gens lab"
            lower_text = text.lower()
            if "sng" in lower_text or "gen's lab" in lower_text or "gens lab" in lower_text:
                print(f"[+] Found match in '{pdf_file}' page {page_num+1}:")
                lines = text.split("\n")
                for line in lines:
                    if any(term in line.lower() for term in ["sng", "gen's lab", "gens lab"]):
                        print(f"    - {line.strip()}")
                found_in_doc = True
        if not found_in_doc:
            pass
    except Exception as e:
        print(f"[!] Error reading '{pdf_file}': {e}")

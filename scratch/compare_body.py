import json
import fitz
import sys
import re

sys.stdout.reconfigure(encoding='utf-8')

json_path = r'C:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main\extracted_jsons\TestReport_AGATAMUDI VENKATA KAMANA (RJM2604120002)_2600130362_526f942c-1bc7-40ec-975b-0c41833042f4.json'
pdf_path = r'C:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main\outsourcing pdf\TestReport_AGATAMUDI VENKATA KAMANA (RJM2604120002)_2600130362_526f942c-1bc7-40ec-975b-0c41833042f4.pdf'

with open(json_path, 'r', encoding='utf-8') as f:
    json_data = json.load(f)

doc = fitz.open(pdf_path)

header_footer_patterns = [
    "QR Code for", "report verification", "Case ID", "Sample Type", "Name : MR.", "AGATAMUDI VENKATA KAMANA",
    "RJM2604120002", "Date & Time Collected", "Sex/Age", "Date & Time Received", "Bill. Loc.",
    "OncQuest Laboratories Ltd", "Date & Time Reported", "Ref. By", "Report Version", "Reviewed by",
    "PRAKASH PATEL", "Dr. Salil Vaniawala", "Dr.Nirmal A. Vaniawala", "M.Sc Ph.D.", "Consulting Geneticist",
    "MD (Path. & Bact.)", "Brainseq Gen"
]

def is_header_footer(line):
    for pat in header_footer_patterns:
        if pat.lower() in line.lower():
            return True
    return False

def extract_strings_from_json(obj):
    res = []
    if isinstance(obj, str):
        if not obj.startswith('data:image'):
            res.append(obj)
    elif isinstance(obj, list):
        for item in obj:
            res.extend(extract_strings_from_json(item))
    elif isinstance(obj, dict):
        for k, v in obj.items():
            if k not in ['data_uri', 'bbox']:
                res.extend(extract_strings_from_json(v))
    return res

print("="*80)
print("BODY CONTENT COMPARISON (EXCLUDING REPETITIVE HEADERS/FOOTERS)")
print("="*80)

for p_idx in range(len(doc)):
    page_num = p_idx + 1
    pdf_page_text = doc[p_idx].get_text()
    
    # Filter out header and footer lines
    raw_lines = [l.strip() for l in pdf_page_text.split('\n') if l.strip()]
    body_lines = [l for l in raw_lines if not is_header_footer(l)]
    
    page_sections = [s for s in json_data.get('sections', []) if s.get('page') == page_num]
    json_page_text = " ".join(extract_strings_from_json(page_sections))
    
    missing_body_lines = []
    for line in body_lines:
        # Check if line is present in json page text
        # Ignore space differences
        norm_line = re.sub(r'\s+', ' ', line)
        norm_json = re.sub(r'\s+', ' ', json_page_text)
        if norm_line not in norm_json:
            missing_body_lines.append(line)
            
    print(f"\n--- PAGE {page_num} ---")
    print(f"Total Body Lines in PDF: {len(body_lines)} | JSON sections on page: {len(page_sections)}")
    print(f"Body Lines Missing from JSON: {len(missing_body_lines)}")
    if missing_body_lines:
        print("  Missing Body Lines:")
        for mbl in missing_body_lines[:20]:
            print(f"    [MISSING]: '{mbl}'")
            
    # Also check extra content in JSON not in PDF
    # (e.g. trailing underscores like IDH1 NM 005896.4 _)
    if "_" in json_page_text:
        trailing_us = re.findall(r'\b\w+\s*_', json_page_text)
        if trailing_us:
            print(f"  Trailing Underscores in JSON Page {page_num}: {set(trailing_us)}")


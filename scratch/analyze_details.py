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

print("="*80)
print("1. HEADER AND FOOTER PATIENT DATA COMPARISON")
print("="*80)

json_str = json.dumps(json_data, ensure_ascii=False)

header_footer_fields = {
    "Case ID": "2600130362",
    "Sample Type": "FFPE TISSUE BLOCK",
    "Patient Name": "MR. AGATAMUDI VENKATA KAMANA",
    "Patient Reg No": "RJM2604120002",
    "Sex/Age": "Male/61 Years",
    "Date & Time Collected": "12-Apr-2026 12:00 AM",
    "Date & Time Received": "15-Apr-2026 04:44 PM",
    "Bill. Loc.": "OncQuest Laboratories Ltd., New Delhi",
    "Date & Time Reported": "28-Apr-2026 06:16 PM",
    "Doctor/Reviewer 1": "PRAKASH PATEL",
    "Doctor/Reviewer 2": "Dr. Salil Vaniawala",
    "Doctor/Reviewer 3": "Dr.Nirmal A. Vaniawala",
    "Qualifications": "Ph.D.(Human Genetics)"
}

for field, val in header_footer_fields.items():
    found_exact = val in json_str
    found_case_insensitive = val.lower() in json_str.lower()
    print(f"[{'EXACT MATCH' if found_exact else ('PARTIAL MATCH' if found_case_insensitive else 'MISSING')}] {field}: '{val}'")

print("\n" + "="*80)
print("2. METADATA & EXTRACTED KEY VALUE PAIRS IN JSON")
print("="*80)
print("Metadata in JSON:", json_data.get('metadata'))
print("Extracted Key-Value Pairs:", json_data.get('extracted_key_value_pairs'))

print("\n" + "="*80)
print("3. PAGE BY PAGE CONTENT DIFFERENCES & MISSING TEXT IN JSON")
print("="*80)

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

for page_idx in range(len(doc)):
    page_num = page_idx + 1
    pdf_text = doc[page_idx].get_text().strip()
    
    page_sections = [s for s in json_data.get('sections', []) if s.get('page') == page_num]
    json_page_text = " ".join(extract_strings_from_json(page_sections)).strip()
    
    print(f"\n--- PAGE {page_num} ---")
    print(f"PDF raw character count: {len(pdf_text)} | JSON sections count: {len(page_sections)} | JSON char count: {len(json_page_text)}")
    
    # Check for text in PDF that is missing from JSON page sections
    lines = [line.strip() for line in pdf_text.split('\n') if line.strip()]
    missing_lines = []
    for line in lines:
        # Check if line or significant part of line is in JSON page text or whole JSON
        clean_line = line.replace('\xa0', ' ')
        if clean_line not in json_page_text and clean_line not in json_str:
            # Ignore standard repetitive header/footer if known, but log it
            missing_lines.append(line)
            
    print(f"Total lines in PDF: {len(lines)}")
    print(f"Lines not present in JSON: {len(missing_lines)}")
    if missing_lines:
        print("  Sample missing lines from PDF:")
        for ml in missing_lines[:15]:
            print(f"    - {ml}")


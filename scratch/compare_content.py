import json
import fitz  # PyMuPDF
import re
from collections import Counter

json_path = r'C:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main\extracted_jsons\TestReport_AGATAMUDI VENKATA KAMANA (RJM2604120002)_2600130362_526f942c-1bc7-40ec-975b-0c41833042f4.json'
pdf_path = r'C:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main\outsourcing pdf\TestReport_AGATAMUDI VENKATA KAMANA (RJM2604120002)_2600130362_526f942c-1bc7-40ec-975b-0c41833042f4.pdf'

with open(json_path, 'r', encoding='utf-8') as f:
    json_data = json.load(f)

doc = fitz.open(pdf_path)

print("="*80)
print("1. GENERAL DOCUMENT SUMMARY COMPARISON")
print("="*80)
print(f"PDF Total Pages: {len(doc)}")
print(f"JSON Total Pages in Summary: {json_data.get('document_summary', {}).get('total_pages')}")
print(f"JSON Total Tables in Summary: {json_data.get('document_summary', {}).get('total_tables')}")
print(f"JSON Total Sections: {len(json_data.get('sections', []))}")

# Helper to flatten JSON text values
def get_all_json_strings(obj):
    strings = []
    if isinstance(obj, str):
        # Ignore data URI base64 images
        if not obj.startswith('data:image'):
            strings.append(obj)
    elif isinstance(obj, list):
        for item in obj:
            strings.extend(get_all_json_strings(item))
    elif isinstance(obj, dict):
        for k, v in obj.items():
            if k not in ['data_uri', 'bbox', 'file_name']:
                strings.extend(get_all_json_strings(v))
    return strings

all_json_strings = get_all_json_strings(json_data)
full_json_raw_text = " ".join(all_json_strings)

# Collect PDF text page by page
pdf_page_texts = [page.get_text() for page in doc]
full_pdf_raw_text = "\n".join(pdf_page_texts)

# Tokenize into words
def clean_tokens(text):
    # Split into words keeping alphanumeric and key symbols
    return re.findall(r'[A-Za-z0-9_\-\.\:\>\<\+\%\(\)\/]+', text)

pdf_tokens = clean_tokens(full_pdf_raw_text)
json_tokens = clean_tokens(full_json_raw_text)

pdf_token_counts = Counter([t.lower() for t in pdf_tokens])
json_token_counts = Counter([t.lower() for t in json_tokens])

print("\n" + "="*80)
print("2. GLOBAL WORD/TOKEN STATS")
print("="*80)
print(f"PDF total word tokens extracted: {len(pdf_tokens)} (Unique: {len(pdf_token_counts)})")
print(f"JSON total word tokens extracted: {len(json_tokens)} (Unique: {len(json_token_counts)})")

# Find tokens present in PDF but completely missing in JSON
missing_in_json = set(pdf_token_counts.keys()) - set(json_token_counts.keys())
# Filter out purely single punctuation or trivial short noise if needed, but keep significant terms
significant_missing = [t for t in missing_in_json if len(t) > 1 and not t.isdigit()]

print(f"\nUnique tokens in PDF missing entirely from JSON: {len(missing_in_json)}")
print(f"Significant text tokens (>1 char, non-pure-digits) missing in JSON ({len(significant_missing)}):")
print(sorted(significant_missing)[:100])

# Page by Page Breakdown
print("\n" + "="*80)
print("3. PAGE BY PAGE DETAILED COMPARISON")
print("="*80)

for p_idx, page_text in enumerate(pdf_page_texts):
    page_num = p_idx + 1
    p_pdf_tokens = clean_tokens(page_text)
    
    # Get JSON strings for this page
    page_sections = [s for s in json_data.get('sections', []) if s.get('page') == page_num]
    page_json_strings = get_all_json_strings(page_sections)
    page_json_text = " ".join(page_json_strings)
    p_json_tokens = clean_tokens(page_json_text)
    
    p_pdf_set = set([t.lower() for t in p_pdf_tokens])
    p_json_set = set([t.lower() for t in p_json_tokens])
    global_json_set = set(json_token_counts.keys())
    
    missing_on_page = p_pdf_set - p_json_set
    missing_globally = p_pdf_set - global_json_set
    
    sig_missing_page = sorted([t for t in missing_on_page if len(t) > 1 and not t.isdigit()])
    sig_missing_glob = sorted([t for t in missing_globally if len(t) > 1 and not t.isdigit()])
    
    print(f"\n--- PAGE {page_num} ---")
    print(f"PDF words: {len(p_pdf_tokens)} | Page JSON words: {len(p_json_tokens)}")
    print(f"Tokens on page missing from Page JSON: {len(missing_on_page)} (Significant: {len(sig_missing_page)})")
    print(f"Tokens on page missing from Entire JSON: {len(missing_globally)} (Significant: {len(sig_missing_glob)})")
    if sig_missing_glob:
        print(f"  Missing globally from JSON on page {page_num}: {sig_missing_glob[:30]}")


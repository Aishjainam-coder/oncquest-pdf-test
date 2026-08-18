import fitz
import re
import difflib

# Dynamic header/footer keywords to filter them out of body content
header_kw = [
    'neuberg', 'center for genomic medicine', 'genomic medicine', 'neuberg diagnostics',
    'laboratory report', 'mc-6200', 'case id', 'sample type', 'patient name', 'uhid', 'reg no',
    'ref by', 'referred by', 'dis.loc.', 'pt id', 'pt. id', 'pt. loc.', 'ph #', 'ref id', 'ref id 2',
    'age/gender', 'age / sex', 'date & time collected', 'date & time received', 'date & time reported',
    'registration date & time', 'sample date & time', 'report date & time',
    'date collected', 'date received', 'date reported', 'report version', 'bill. loc.',
    'lab no', 'barcode', 'qr code for report', 'report verification', 'patient information',
    'patient details', 'patient metadata', 'demographics', 'sample coll.by', 'acc. remarks',
    'mr. dulal naha', 'kol2604070057', 'blood in ccf dna tube', '2600128556', 'oncquest laboratories',
    'new delhi'
]

footer_kw = [
    'reviewed by', 'verified by', 'authorized signatory', 'signatory', 'doctor', 'dr.',
    'md (path', 'ph.d.', 'pathologist', 'biochemist', 'microbiologist', 'geneticist',
    'consultant', 'end of report', 'electronically generated', 'disclaimer', 'registered office',
    'nabl', 'cap accredited', 'iso 15189', 'mc-7414', 'page ', 'dr.nirmal a. vaniawala', 
    'dr. salil vaniawala', 'consulting geneticist'
]

def clean_and_normalize_text(text):
    # Split text into lines, filter out lines containing header/footer keywords
    cleaned_lines = []
    for line in text.splitlines():
        line_s = line.strip()
        if not line_s:
            continue
        line_l = line_s.lower()
        # Skip header or footer lines
        if any(kw in line_l for kw in header_kw) or any(kw in line_l for kw in footer_kw):
            continue
        # Skip standalone page numbers or NABL
        if re.match(r'^page\s+\d+\s+of\s+\d+$', line_l):
            continue
        # Normalize whitespace
        norm_line = re.sub(r'\s+', ' ', line_s)
        cleaned_lines.append(norm_line)
        
    return cleaned_lines

def compare_text_strictly(orig_pdf_path, comp_pdf_path):
    doc_orig = fitz.open(orig_pdf_path)
    doc_comp = fitz.open(comp_pdf_path)
    
    print(f"Comparing text strictly (ignoring header/footer/page numbers)...")
    for i in range(len(doc_orig)):
        page_num = i + 1
        orig_text = doc_orig[i].get_text("text")
        comp_text = doc_comp[i].get_text("text")
        
        orig_lines = clean_and_normalize_text(orig_text)
        comp_lines = clean_and_normalize_text(comp_text)
        
        # We also want to compare as full text blocks (words) to see if anything is different
        orig_full = " ".join(orig_lines)
        comp_full = " ".join(comp_lines)
        
        # Normalize and remove punctuation/whitespace for strict character check
        orig_norm = re.sub(r'\s+', '', orig_full).lower()
        comp_norm = re.sub(r'\s+', '', comp_full).lower()
        
        print(f"\n--- Page {page_num} Strict Compare ---")
        if orig_norm == comp_norm:
            print(f"  Status: TEXT CONTENT MATCHES EXACTLY (ignoring formatting/newlines)")
        else:
            print(f"  Status: TEXT CONTENT MISMATCH")
            # Generate diff of the lines
            diff = list(difflib.unified_diff(
                orig_lines,
                comp_lines,
                fromfile=f"Orig_Page_{page_num}",
                tofile=f"Comp_Page_{page_num}",
                lineterm=""
            ))
            for line in diff[:30]:  # Show first 30 diff lines
                print("   ", line)
            if len(diff) > 30:
                print(f"    ... and {len(diff) - 30} more diff lines.")
                
    doc_orig.close()
    doc_comp.close()

if __name__ == "__main__":
    orig = r"outsourcing pdf/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.pdf"
    comp = r"output/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee_compiled.pdf"
    compare_text_strictly(orig, comp)

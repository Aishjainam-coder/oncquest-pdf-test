import re
import copy
import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pathlib import Path
import sys

sys.stdout.reconfigure(encoding='utf-8')

# Strict Test Name Patterns
STANDALONE_TEST_PATTERNS = [
    r'^(?:Breast\s+and\s+Ovarian\s+Extended\s+Panel\s*[-–]\s*Liquid\s+Biopsy\s+Assay)$',
    r'^(?:(?:Liquidseq\s+Actionable|Brainseq)\s+Genomic\s+Profiling\s+Panel(?:\s*[-–]\s*Advance)?)$',
    r'^(?:Whole\s+Exome\s+Sequencing(?:\s+on\s+(?:the\s+)?Illumina\s+[\w\s-]+\s+Platform)?)$',
    r'^[\w\s/&,–\-\(\)\.\+]+?(?:Genomic\s+Profiling\s+Panel|Extended\s+Panel|Profiling\s+Panel|Biopsy\s+Assay|Exome\s+Sequencing|Sequencing\s+Panel|Cancer\s+Panel|Gene\s+Panel|Profiling\s+Assay|Sequencing\s+Assay|Biopsy\s+Panel|NGS\s+Panel)(?:\s*[-–]\s*Advance)?(?:\s*\([^)]*\))?(?:\s+on\s+(?:the\s+)?Illumina\s+[\w\s-]+\s+Platform)?$',
    r'^(?:[A-Z\s]{4,}\s+PANEL(?:\s*[-–]\s*ADVANCE)?(?:\s*\([^)]*\))?)$',
]

def is_test_name_text(text: str) -> bool:
    if not isinstance(text, str):
        return False
    cleaned = text.strip()
    if len(cleaned) < 4 or len(cleaned) > 130:
        return False
        
    exclude_prefixes = (
        "clinical indication", "sample description", "report highlights", 
        "key findings", "test results", "tier ", "case id", "sample type", 
        "name :", "date & time", "bill. loc", "ref. by", "report version",
        "qr code", "page ", "salient features", "clinical suspicion",
        "dr.", "laboratory", "oncquest", "result summary", "methodology",
        "segregation analysis", "this assay", "the assay", "note:", "interpretation"
    )
    cleaned_lower = cleaned.lower()
    for ex in exclude_prefixes:
        if cleaned_lower.startswith(ex):
            return False
            
    for pat in STANDALONE_TEST_PATTERNS:
        if re.match(pat, cleaned, re.IGNORECASE):
            return True
            
    return False

def is_end_of_report_text(text: str) -> bool:
    if not isinstance(text, str):
        return False
    return bool(re.search(r'end\s+of\s+report', text, re.IGNORECASE))

def set_run_highlight_yellow(r_elem):
    rPr = r_elem.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        r_elem.insert(0, rPr)
    
    # Highlight yellow
    hl = rPr.find(qn('w:highlight'))
    if hl is None:
        hl = OxmlElement('w:highlight')
        rPr.append(hl)
    hl.set(qn('w:val'), 'yellow')
    
    # Ensure text is black/dark so it is readable on yellow highlight
    color = rPr.find(qn('w:color'))
    if color is not None:
        c_val = (color.get(qn('w:val')) or "").lower()
        if c_val in ('ffffff', 'fff', 'yellow', 'auto'):
            color.set(qn('w:val'), '000000')

def highlight_test_name_in_paragraph(p_elem, seen_eor=False):
    """
    Highlights test name occurrences in a <w:p> XML element:
    1. Exact 'TEST NAME' substrings
    2. 'Test Name: <val>' metadata
    3. Standalone test names matching is_test_name_text()
    """
    runs = [r for r in p_elem if r.tag.endswith('}r')]
    if not runs:
        return 0

    full_p_text = ""
    run_spans = [] # (run, t_elems, r_text, start, end)
    for r in runs:
        t_elems = [t for t in r if t.tag.endswith('}t')]
        r_text = "".join(t.text for t in t_elems if t.text)
        start = len(full_p_text)
        end = start + len(r_text)
        full_p_text += r_text
        run_spans.append((r, t_elems, r_text, start, end))

    if not full_p_text.strip():
        return 0

    cleaned = full_p_text.strip()
    count = 0

    # 1. Standalone test name paragraph (e.g. Title Banner)
    if not seen_eor and (cleaned.upper() == "TEST NAME" or is_test_name_text(cleaned)):
        for r in runs:
            set_run_highlight_yellow(r)
            count += 1
        return count

    # 2. Target specific "TEST NAME" words/phrases within paragraph
    target_spans = []
    
    # Match "TEST NAME" (case insensitive)
    for m in re.finditer(r'\bTEST\s+NAME\b', full_p_text, re.IGNORECASE):
        target_spans.append((m.start(), m.end(), m.group(0)))

    # Match "Test Name\s*:\s*(.+)" in metadata lines
    m_meta = re.search(r'(?:Test\s+Name\s*:\s*)([^\r\n\t]+)', full_p_text, re.IGNORECASE)
    if m_meta:
        target_spans.append((m_meta.start(1), m_meta.end(1), m_meta.group(1)))

    if not target_spans:
        return 0

    # Apply highlight to runs matching target spans
    for m_start, m_end, m_text in target_spans:
        for r, t_elems, r_text, r_start, r_end in run_spans:
            if r_end <= m_start or r_start >= m_end:
                continue
            
            # If the run is completely inside match:
            if r_start >= m_start and r_end <= m_end:
                set_run_highlight_yellow(r)
                count += 1
            else:
                # Run partially contains the match
                if m_text in r_text and len(t_elems) > 0:
                    idx = r_text.find(m_text)
                    before_text = r_text[:idx]
                    after_text = r_text[idx + len(m_text):]
                    
                    t_elems[0].text = before_text
                    
                    rPr_orig = r.find(qn('w:rPr'))
                    new_r = OxmlElement('w:r')
                    if rPr_orig is not None:
                        new_rPr = copy.deepcopy(rPr_orig)
                    else:
                        new_rPr = OxmlElement('w:rPr')
                    new_r.append(new_rPr)
                    
                    hl = new_rPr.find(qn('w:highlight'))
                    if hl is None:
                        hl = OxmlElement('w:highlight')
                        new_rPr.append(hl)
                    hl.set(qn('w:val'), 'yellow')
                    
                    color = new_rPr.find(qn('w:color'))
                    if color is not None:
                        c_val = (color.get(qn('w:val')) or "").lower()
                        if c_val in ('ffffff', 'fff', 'yellow', 'auto'):
                            color.set(qn('w:val'), '000000')

                    new_t = OxmlElement('w:t')
                    new_t.text = m_text
                    new_t.set(qn('xml:space'), 'preserve')
                    new_r.append(new_t)
                    
                    p_elem.insert(p_elem.index(r) + 1, new_r)
                    
                    if after_text:
                        after_r = OxmlElement('w:r')
                        if rPr_orig is not None:
                            after_r.append(copy.deepcopy(rPr_orig))
                        after_t = OxmlElement('w:t')
                        after_t.text = after_text
                        after_t.set(qn('xml:space'), 'preserve')
                        after_r.append(after_t)
                        p_elem.insert(p_elem.index(new_r) + 1, after_r)
                        
                    count += 1

    return count

def highlight_all_test_names_in_doc(doc):
    total = 0
    seen_eor = False
    
    # Process document body
    for p in doc.element.iter():
        if p.tag.endswith('}p'):
            p_text = "".join(t.text for t in p.iter() if t.tag.endswith('}t') and t.text)
            if is_end_of_report_text(p_text):
                seen_eor = True
            total += highlight_test_name_in_paragraph(p, seen_eor=seen_eor)
            
    # Process headers and footers
    for section in doc.sections:
        for h in (section.header, section.first_page_header, section.even_page_header):
            if h is not None:
                for p in h._element.iter():
                    if p.tag.endswith('}p'):
                        total += highlight_test_name_in_paragraph(p, seen_eor=False)
        for f in (section.footer, section.first_page_footer, section.even_page_footer):
            if f is not None:
                for p in f._element.iter():
                    if p.tag.endswith('}p'):
                        total += highlight_test_name_in_paragraph(p, seen_eor=False)
    return total

# Test across multiple files
for path in ['output/TestReport_BIRENDRA KR TRIPATHI (KOL2604200324)_2600132470_e5d40a68-7e61-43d6-aae0-577a57bb598a.docx',
             'output/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.docx',
             'output/TestReport_SAROJ DEVI (OQG2604250052)_2600133939_1eb3f888-f97e-4181-b84d-d9536de7af26.docx']:
    try:
        doc = docx.Document(path)
        cnt = highlight_all_test_names_in_doc(doc)
        print(f"Highlighted {cnt} in {Path(path).name}")
    except Exception as e:
        print(f"Error {path}: {e}")

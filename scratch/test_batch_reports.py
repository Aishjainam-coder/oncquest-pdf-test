import glob
import json
import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import docx
from docx.oxml.ns import qn
from converter import convert_json_to_docx, generate_dynamic_template_html

def test_all_reports():
    json_files = glob.glob('extracted_jsons/*.json')
    print(f"Testing {len(json_files)} extracted JSON files...")
    
    for jf in json_files:
        fname = os.path.basename(jf)
        try:
            with open(jf, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            docx_bytes = convert_json_to_docx(data)
            out_path = f"scratch/test_out_{fname[:20]}.docx"
            with open(out_path, 'wb') as f:
                f.write(docx_bytes)
                
            doc = docx.Document(out_path)
            hl_list = []
            for p in doc.paragraphs:
                for r in p.runs:
                    rPr = r._r.find(qn('w:rPr'))
                    if rPr is not None:
                        hl = rPr.find(qn('w:highlight'))
                        if hl is not None:
                            hl_list.append((r.text, hl.get(qn('w:val'))))
            print(f"[OK] {fname}: {len(hl_list)} highlighted runs -> {hl_list[:3]}")
        except Exception as e:
            print(f"[FAIL] {fname}: {e}")

if __name__ == "__main__":
    test_all_reports()

import glob, json, os

for jf in glob.glob('extracted_jsons/*.json'):
    try:
        with open(jf, encoding='utf-8') as f:
            d = json.load(f)
        pages = d.get('document', {}).get('pages', []) or d.get('pages', [])
        for p in pages:
            elems = p.get('elements', [])
            for idx, el in enumerate(elems):
                txt = str(el.get('text', '')) + str(el.get('data', '')) + str(el.get('title', ''))
                if 'TEST NAME' in txt.upper() or 'SOLIDSEQ' in txt.upper() or 'LIQUIDSEQ' in txt.upper() or 'WHOLE EXOME' in txt.upper():
                    next_type = elems[idx+1].get('type') if idx+1 < len(elems) else 'NONE'
                    next_txt = str(elems[idx+1].get('text', ''))[:40] if idx+1 < len(elems) else ''
                    print(f"{os.path.basename(jf)}: TN at {idx} -> Next elem: {next_type} {repr(next_txt)}")
    except Exception as e:
        print(f"Error on {jf}: {e}")

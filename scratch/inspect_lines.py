import fitz, glob, os

for pdf in glob.glob('outsourcing pdf/*.pdf')[:5]:
    print('=== PDF:', os.path.basename(pdf))
    doc = fitz.open(pdf)
    page = doc[0]
    p_dict = page.get_text('dict')
    tn_boxes = []
    for b in p_dict.get('blocks', []):
        for l in b.get('lines', []):
            txt = ''.join(s.get('text', '') for s in l.get('spans', []))
            if any(k in txt.upper() for k in ['SOLIDSEQ', 'LIQUIDSEQ', 'WHOLE EXOME', 'PANEL', 'TEST NAME']):
                print(f"  TN text: '{txt.strip()}' bbox: {l['bbox']}")
                tn_boxes.append((l['bbox'][1], l['bbox'][3]))
    
    drawings = page.get_drawings()
    print(f"  Total drawings on page 1: {len(drawings)}")
    for d in drawings:
        r = d['rect']
        rw, rh = r.x1 - r.x0, r.y1 - r.y0
        if rh <= 3.0 and rw > 10.0:
            for y0, y1 in tn_boxes:
                if (y0 - 50) <= r.y0 <= (y1 + 50):
                    print(f"    Line near TN: y0={r.y0:.1f}, y1={r.y1:.1f}, w={rw:.1f}, h={rh:.1f}, dist_from_top={r.y0-y0:.1f}, dist_from_bottom={r.y0-y1:.1f}")

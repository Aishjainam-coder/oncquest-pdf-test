import re

def parse_blocks_file(filepath):
    pages = {}
    current_page = None
    with open(filepath, "r", encoding="utf-8") as f:
        content = f.read()
        
    page_sections = content.split("=========================================\nPAGE ")
    for sec in page_sections[1:]:
        lines = sec.splitlines()
        page_num = int(lines[0].strip().split()[0])
        pages[page_num] = []
        
        block_text = None
        bbox = None
        for line in lines[2:]:
            if line.startswith("  Block "):
                # Parse bbox
                m = re.search(r"bbox: \[(.*?)\], type: (\d+)", line)
                if m:
                    bbox = [float(val) for val in m.group(1).split(",")]
            elif line.startswith("    \""):
                block_text = line.strip().strip('"')
                pages[page_num].append({
                    "bbox": bbox,
                    "text": block_text
                })
    return pages

def is_header_footer(text):
    text_l = text.lower()
    # Check for header fields
    header_terms = [
        "case id", "sample type", "date & time", "sex/age", "bill. loc.", "ref. by", 
        "report version", "indication", "qr code", "dulal naha", "kol2604070057",
        "blood in ccf dna", "oncquest laboratories", "new delhi", "report verification"
    ]
    # Check for footer fields (signatures/doctors/page numbers)
    footer_terms = [
        "dr.nirmal", "dr. salil", "vaniawala", "md (path", "consulting geneticist", "human genetics",
        "page 1 of", "page 2 of", "page 3 of", "page 4 of", "page 5 of", 
        "page 6 of", "page 7 of", "page 8 of", "page 9 of", "page 10 of",
        "end of report"
    ]
    
    # Check if text matches header/footer terms
    if any(term in text_l for term in header_terms) or any(term in text_l for term in footer_terms):
        return True
    return False

def compare_pages(orig_pages, comp_pages):
    all_pages = sorted(list(set(orig_pages.keys()) | set(comp_pages.keys())))
    
    for page in all_pages:
        print(f"\n=========================================")
        print(f"PAGE {page} COMPARISON")
        print(f"=========================================")
        
        orig_list = orig_pages.get(page, [])
        comp_list = comp_pages.get(page, [])
        
        print(f"Original blocks total: {len(orig_list)}")
        print(f"Compiled blocks total: {len(comp_list)}")
        
        # Categorize original blocks
        orig_content_blocks = []
        orig_header_footer_blocks = []
        for ob in orig_list:
            if is_header_footer(ob["text"]):
                orig_header_footer_blocks.append(ob)
            else:
                orig_content_blocks.append(ob)
                
        # Categorize compiled blocks
        comp_content_blocks = []
        comp_header_footer_blocks = []
        for cb in comp_list:
            if is_header_footer(cb["text"]):
                comp_header_footer_blocks.append(cb)
            else:
                comp_content_blocks.append(cb)
                
        print(f"Original: {len(orig_header_footer_blocks)} header/footers, {len(orig_content_blocks)} content blocks")
        print(f"Compiled: {len(comp_header_footer_blocks)} header/footers, {len(comp_content_blocks)} content blocks")
        
        # Match content blocks
        matched_comp_indices = set()
        mismatches = []
        missing = []
        exact_matches = []
        
        for ob in orig_content_blocks:
            best_match_idx = -1
            best_match_score = 0.0
            
            # Simple matching: check if text is very similar
            for idx, cb in enumerate(comp_content_blocks):
                if idx in matched_comp_indices:
                    continue
                # Exact or substring match
                if ob["text"] == cb["text"]:
                    best_match_idx = idx
                    best_match_score = 1.0
                    break
                # Let's check similarity by overlap
                words_ob = set(ob["text"].lower().split())
                words_cb = set(cb["text"].lower().split())
                if words_ob and words_cb:
                    overlap = len(words_ob & words_cb) / max(len(words_ob), len(words_cb))
                    if overlap > best_match_score:
                        best_match_score = overlap
                        best_match_idx = idx
            
            if best_match_score == 1.0:
                exact_matches.append((ob, comp_content_blocks[best_match_idx]))
                matched_comp_indices.add(best_match_idx)
            elif best_match_score > 0.4:
                mismatches.append((ob, comp_content_blocks[best_match_idx], best_match_score))
                matched_comp_indices.add(best_match_idx)
            else:
                missing.append(ob)
                
        extra = [cb for idx, cb in enumerate(comp_content_blocks) if idx not in matched_comp_indices]
        
        print("\n--- EXACT MATCHES ---")
        print(f"Count: {len(exact_matches)}")
        
        if mismatches:
            print("\n--- MISMATCHED / REORGANIZED CONTENT ---")
            for ob, cb, score in mismatches:
                print(f"  Similarity: {score:.2f}")
                print(f"  Original: \"{ob['text']}\"")
                print(f"  Compiled: \"{cb['text']}\"")
                
        if missing:
            print("\n--- MISSING CONTENT ---")
            for ob in missing:
                print(f"  - \"{ob['text']}\"")
                
        if extra:
            print("\n--- EXTRA / ADDED / DUPLICATED CONTENT ---")
            for cb in extra:
                print(f"  - \"{cb['text']}\"")
                
        if orig_header_footer_blocks:
            print("\n--- HEADER/FOOTER (Removed from Compiled body as expected) ---")
            print(f"  Removed {len(orig_header_footer_blocks)} header/footer block(s).")
            # Print sample to verify
            for ob in orig_header_footer_blocks[:2]:
                print(f"    - \"{ob['text']}\"")
            if len(orig_header_footer_blocks) > 2:
                print(f"    ... and {len(orig_header_footer_blocks)-2} more.")

if __name__ == "__main__":
    orig_pages = parse_blocks_file("scratch/orig_blocks.txt")
    comp_pages = parse_blocks_file("scratch/comp_blocks.txt")
    compare_pages(orig_pages, comp_pages)

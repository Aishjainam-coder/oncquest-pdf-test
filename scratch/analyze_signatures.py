import re

html_path = "output/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee_template.html"

with open(html_path, "r", encoding="utf-8") as f:
    html = f.read()

# Let's find all page divs
pages = re.findall(r"<div class='pdf-page'[^>]*>.*?</div>\s*(?=<div class='pdf-page'|</div></body></html>|$)", html, re.DOTALL)
print(f"Total pages in HTML: {len(pages)}")

for idx, page in enumerate(pages):
    print(f"\n--- PAGE {idx+1} ---")
    # Find all images on this page
    imgs = re.findall(r"<img[^>]*>", page)
    print(f"Found {len(imgs)} images on this page:")
    for img_idx, img in enumerate(imgs):
        # Extract style and base64 prefix
        style = re.search(r"style='([^']*)'|style=\"([^\"]*)\"", img)
        style_val = style.group(1) or style.group(2) if style else "No style"
        src_match = re.search(r"src='data:image/([^;]*);base64,([^']*)'|src=\"data:image/([^;]*);base64,([^\"]*)\"", img)
        if src_match:
            img_type = src_match.group(1) or src_match.group(3)
            b64_len = len(src_match.group(2) or src_match.group(4))
            print(f"  Img {img_idx+1}: type={img_type}, b64_len={b64_len}, style={style_val}")
        else:
            print(f"  Img {img_idx+1}: style={style_val} (non-b64 src)")

from bs4 import BeautifulSoup

def analyze_html(file_path):
    with open(file_path, "r", encoding="utf-8") as f:
        soup = BeautifulSoup(f.read(), "html.parser")
        
    pages = soup.find_all(id=lambda x: x and x.startswith("page-"))
    print(f"Total pages in HTML: {len(pages)}")
    
    for i, page in enumerate(pages):
        print(f"\n--- Page {i+1} ---")
        imgs = page.find_all("img")
        print(f"Total images on page: {len(imgs)}")
        for img_idx, img in enumerate(imgs):
            cls = img.get("class", "None")
            style = img.get("style", "None")
            src = img.get("src", "")
            print(f"  Image {img_idx+1}:")
            print(f"    Class: {cls}")
            print(f"    Style: {style}")
            print(f"    Src: {src[:60]}... (length: {len(src)})")

if __name__ == "__main__":
    analyze_html("output/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee_template.html")

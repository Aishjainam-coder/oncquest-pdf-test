import fitz
from pathlib import Path

def extract_images(pdf_path, out_dir):
    doc = fitz.open(pdf_path)
    out_path = Path(out_dir)
    out_path.mkdir(exist_ok=True)
    
    extracted_xrefs = set()
    
    for page_idx in range(len(doc)):
        page = doc[page_idx]
        image_list = page.get_images(full=True)
        for img in image_list:
            xref = img[0]
            if xref in extracted_xrefs:
                continue
            extracted_xrefs.add(xref)
            
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            image_ext = base_image["ext"]
            
            filename = out_path / f"orig_img_{xref}.{image_ext}"
            with open(filename, "wb") as f:
                f.write(image_bytes)
            print(f"Extracted image xref {xref} to {filename}")
            
    doc.close()

if __name__ == "__main__":
    extract_images(r"outsourcing pdf/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.pdf", "output/orig_images")

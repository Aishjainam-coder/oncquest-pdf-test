import fitz

def check_pdf_images(pdf_path):
    doc = fitz.open(pdf_path)
    print(f"Checking images in PDF: {pdf_path}")
    print(f"Total pages: {len(doc)}")
    
    for page_idx in range(len(doc)):
        page = doc[page_idx]
        image_list = page.get_images(full=True)
        print(f"Page {page_idx + 1} has {len(image_list)} image(s)")
        for img_idx, img in enumerate(image_list):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            image_ext = base_image["ext"]
            print(f"  Image {img_idx + 1}: xref={xref}, ext={image_ext}, size={len(image_bytes)} bytes")
            
        # Let's also check for vector graphics
        paths = page.get_drawings()
        print(f"  Page {page_idx + 1} has {len(paths)} drawing path(s)")

if __name__ == "__main__":
    check_pdf_images(r"outsourcing pdf/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.pdf")

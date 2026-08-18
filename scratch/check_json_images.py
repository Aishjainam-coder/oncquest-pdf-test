import json

def check_json_images(json_path):
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)
        
    print("=== IMAGES AND GRAPHS IN JSON ===")
    img_list = data.get("images_and_graphs", [])
    print(f"Total images extracted in JSON: {len(img_list)}")
    for i, img in enumerate(img_list):
        print(f"\nImage {i+1}:")
        print(f"  Type: {img.get('type')}")
        print(f"  Page Number: {img.get('page_number')}")
        print(f"  BBox: {img.get('bbox')}")
        src = img.get("image_base64", "")
        print(f"  Base64 length: {len(src)}")

if __name__ == "__main__":
    check_json_images("extracted_jsons/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.json")

from PIL import Image
from pathlib import Path

assets_dir = Path("assets")
for p in assets_dir.glob("*"):
    if p.suffix.lower() in [".png", ".jpg", ".jpeg"]:
        try:
            with Image.open(p) as img:
                print(f"{p.name}: format={img.format}, size={img.size}, mode={img.mode}")
        except Exception as e:
            print(f"Error reading {p.name}: {e}")

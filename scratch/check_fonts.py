import os
from pathlib import Path

fonts_dir = Path("C:/Windows/Fonts")
cambria_files = list(fonts_dir.glob("cambria*"))

print("Cambria font files found:")
for f in cambria_files:
    print(f"  {f.name} (size: {f.stat().st_size} bytes)")

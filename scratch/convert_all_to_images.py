import fitz
from pathlib import Path

workspace_dir = Path(r"c:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main")
downloads_dir = Path(r"c:\Users\aishwarya.jain\Downloads")
output_img_dir = workspace_dir / "output_images"
output_img_dir.mkdir(exist_ok=True)

unique_pdfs = [
    workspace_dir / "output/exact_position_test_compiled.pdf",
    workspace_dir / "output/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee_compiled.pdf",
    workspace_dir / "outsourcing pdf/TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.pdf",
    downloads_dir / "TestReport_BASANTI DEVI 10078040 OQG2604170493_60400127778_42f516a3-ebf4-4187-a505-e2b8c4b5e928_target_output.pdf",
    downloads_dir / "TestReport_BIRENDRA KR TRIPATHI (KOL2604200324)_2600132470_e5d40a68-7e61-43d6-aae0-577a57bb598a_target_output.pdf",
    downloads_dir / "TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.pdf",
    downloads_dir / "vendor pdf_output (1).pdf",
    downloads_dir / "vendor pdf_output (2).pdf",
    downloads_dir / "vendor pdf_output (3).pdf",
    downloads_dir / "vendor pdf_output.pdf"
]

for p in unique_pdfs:
    if not p.exists():
        continue
    try:
        doc = fitz.open(p)
        page = doc[0]
        pix = page.get_pixmap(dpi=150)
        out_name = f"{p.stem}_page_1.png"
        pix.save(str(output_img_dir / out_name))
        print(f"Saved: {out_name}")
        doc.close()
    except Exception as e:
        print(f"Error rendering {p.name}: {e}")

import fitz
from pathlib import Path

pdf_paths = [
    Path(r"c:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main\outsourcing pdf\TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee.pdf"),
    Path(r"c:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main\output\TestReport_DULAL NAHA (KOL2604070057)_2600128556_11d3f985-7e08-40be-91b4-48b35be363ee_compiled.pdf"),
    Path(r"c:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main\output\exact_position_test_compiled.pdf"),
    Path(r"c:\Users\aishwarya.jain\Downloads\TestReport_BASANTI DEVI 10078040 OQG2604170493_60400127778_42f516a3-ebf4-4187-a505-e2b8c4b5e928_target_output.pdf"),
    Path(r"c:\Users\aishwarya.jain\Downloads\TestReport_BIRENDRA KR TRIPATHI (KOL2604200324)_2600132470_e5d40a68-7e61-43d6-aae0-577a57bb598a_target_output.pdf"),
    Path(r"c:\Users\aishwarya.jain\Downloads\TestReport_DEVIKA 2908914 OQG2603250447_60300129567_39c65aa4-7a37-4af2-a4e6-46c303a638fe_target_output.pdf"),
    Path(r"c:\Users\aishwarya.jain\Downloads\TestReport_SAROJ DEVI (OQG2604250052)_2600133939_1eb3f888-f97e-4181-b84d-d9536de7af26_output.pdf"),
    Path(r"c:\Users\aishwarya.jain\Downloads\vendor pdf_output.pdf")
]

output_file = Path(r"c:\Users\aishwarya.jain\Downloads\Oncquest-main\Oncquest-main\scratch\spans_inspection.txt")

with open(output_file, "w", encoding="utf-8") as f:
    for p in pdf_paths:
        if not p.exists():
            f.write(f"Path does not exist: {p}\n\n")
            continue
        f.write("="*80 + "\n")
        f.write(f"File: {p.name}\n")
        f.write("="*80 + "\n")
        try:
            doc = fitz.open(p)
            page = doc[0]
            f.write(f"Page width: {page.rect.width}, height: {page.rect.height}\n")
            
            # Use page.get_text("dict") to get structured text with spans
            text_dict = page.get_text("dict")
            for block_idx, block in enumerate(text_dict["blocks"]):
                if "lines" not in block:
                    continue
                f.write(f"  Block {block_idx} (bbox: {block['bbox']}):\n")
                for line_idx, line in enumerate(block["lines"]):
                    f.write(f"    Line {line_idx} (bbox: {line['bbox']}):\n")
                    for span_idx, span in enumerate(line["spans"]):
                        f.write(
                            f"      Span {span_idx}: text='{span['text']}' "
                            f"font='{span['font']}' size={span['size']:.2f} "
                            f"color={span['color']} bbox={span['bbox']} "
                            f"flags={span['flags']}\n"
                        )
            doc.close()
        except Exception as e:
            f.write(f"Error reading {p.name}: {e}\n")
        f.write("\n\n")

print(f"Inspection written to {output_file}")

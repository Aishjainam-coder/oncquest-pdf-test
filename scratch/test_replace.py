import re

def replace_sng_gen_lab(text: str) -> str:
    if not isinstance(text, str):
        return text
    # Matches variations of "SNG Gen's Lab pvt ltd" case-insensitively, handling optional spaces, periods and different apostrophes/HTML entities.
    pattern = re.compile(r"SNG\s+Gene?(?:['’‘]|&[a-zA-Z0-9#]+;)?s\s+Lab\s+pvt\.?\s*ltd", re.IGNORECASE)
    return pattern.sub("Laboratory", text)

# Test cases
test_cases = [
    "SNG Gen&apos;s Lab pvt ltd ",
    "SNG Gen&#39;s Lab pvt. ltd.",
    "SNG Gen&rsquo;s Lab Pvt Ltd",
    "SNG Gen's Lab pvt ltd"
]

print("Starting test...")
for tc in test_cases:
    replaced = replace_sng_gen_lab(tc)
    print(f"Original: {repr(tc)}")
    print(f"Replaced: {repr(replaced)}")
    print("-" * 40)

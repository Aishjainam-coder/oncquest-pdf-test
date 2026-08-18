import fitz

doc = fitz.open()
page = doc.new_page()

print("Page methods matching 'redact':")
for m in dir(page):
    if "redact" in m.lower():
        print(f"  {m}")

print("Document methods matching 'redact':")
for m in dir(doc):
    if "redact" in m.lower():
        print(f"  {m}")
doc.close()

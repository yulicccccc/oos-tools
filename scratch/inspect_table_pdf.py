import os
import fitz

desktop = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
pdf_path = os.path.join(desktop, "Celsis table OOS-262017.pdf")
doc = fitz.open(pdf_path)
print(f"Table PDF pages: {len(doc)}")
for i, p in enumerate(doc):
    print(f"\n--- Page {i+1} ---")
    print(p.get_text())

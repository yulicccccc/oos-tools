import os, sys
import fitz

desktop = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
tables_pdf = os.path.join(desktop, "Celsis table OOS-262017.pdf")
doc_tgt = fitz.open(os.path.join(desktop, "CORP-FORM-21 - P1 31 Aug 2026.pdf"))
doc_tbl = fitz.open(tables_pdf)

doc_tgt.insert_pdf(doc_tbl, links=True)
print(f"Total pages after inserting tables: {len(doc_tgt)}")

p7_links = doc_tgt[6].get_links()
p8_links = doc_tgt[7].get_links()

print(f"Page 7 links count: {len(p7_links)}")
for l in p7_links:
    print("  P7:", l.get('uri'))

print(f"Page 8 links count: {len(p8_links)}")
for l in p8_links:
    print("  P8:", l.get('uri'))

doc_tgt.close()
doc_tbl.close()

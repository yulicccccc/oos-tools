import os
import fitz

desktop = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
target_form = os.path.join(desktop, "CORP-FORM-21 - P1 31 Aug 2026.pdf")
doc = fitz.open(target_form)
for i in range(len(doc)):
    txt = doc[i].get_text()
    lines = [l.strip() for l in txt.split('\n') if l.strip()]
    footer_lines = [l for l in lines if any(k in l.lower() for k in ['printed', 'zenqms', 'page ', 'v11', 'corp-form'])]
    print(f"Page {i+1} footer lines: {footer_lines}")

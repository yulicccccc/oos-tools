import os
import fitz

desktop = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
target_form_path = os.path.join(desktop, "CORP-FORM-21 - P1 31 Aug 2026.pdf")
doc_tgt = fitz.open(target_form_path)

print("Default font sizes in target form CORP-FORM-21:")
for pno in range(len(doc_tgt)):
    for w in doc_tgt[pno].widgets():
        if w.field_type_string == 'Text' and any(k in w.field_name for k in ['Field49', 'Field50', 'Field51', 'Field13', 'Field22', 'Field3', 'Field7', 'Field0']):
            print(f"  P{pno+1} {w.field_name}: text_fontsize={w.text_fontsize}, rect={w.rect}")

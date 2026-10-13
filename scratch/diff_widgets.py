import os
import fitz

desktop = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
target_form_path = os.path.join(desktop, "CORP-FORM-21 - P1 31 Aug 2026.pdf")
user_oos_path = os.path.join(desktop, "OOS-262017 Revive Rx Pharmacy - PO Required (E00927) - Celsis.pdf")

doc_tgt = fitz.open(target_form_path)
doc_src = fitz.open(user_oos_path)

print(f"Target page count: {len(doc_tgt)}, Source page count: {len(doc_src)}")

for pno in range(min(len(doc_tgt), len(doc_src))):
    tgt_widgets = {w.field_name: w for w in doc_tgt[pno].widgets()}
    src_widgets = {w.field_name: w for w in doc_src[pno].widgets()}
    
    print(f"\n--- Page {pno+1} ---")
    print(f"Target widget count: {len(tgt_widgets)}, Source widget count: {len(src_widgets)}")
    
    missing_in_tgt = set(src_widgets.keys()) - set(tgt_widgets.keys())
    missing_in_src = set(tgt_widgets.keys()) - set(src_widgets.keys())
    
    if missing_in_tgt:
        print(f"  In Source but NOT in Target: {missing_in_tgt}")
    if missing_in_src:
        print(f"  In Target but NOT in Source: {missing_in_src}")
        for fn in missing_in_src:
            w = tgt_widgets[fn]
            print(f"    Target field '{fn}': type={w.field_type_string}, rect={w.rect}")

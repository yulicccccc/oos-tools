import fitz

doc_user = fitz.open(r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\CORP-FORM-21 - P1 - 16 Sep 2026.pdf")
doc_tpl = fitz.open(r"ScanRDI OOS template.pdf")

print("User pages:", len(doc_user), "Template pages:", len(doc_tpl))

for pno in range(min(len(doc_user), len(doc_tpl))):
    u_widgets = {w.field_name: (w.field_type, w.rect) for w in doc_user[pno].widgets()}
    t_widgets = {w.field_name: (w.field_type, w.rect) for w in doc_tpl[pno].widgets()}
    
    diff_keys = set(u_widgets.keys()) ^ set(t_widgets.keys())
    print(f"Page {pno+1}: user widgets={len(u_widgets)}, tpl widgets={len(t_widgets)}, symmetric diff={len(diff_keys)}")
    if diff_keys:
        print(f"  Diff keys: {diff_keys}")

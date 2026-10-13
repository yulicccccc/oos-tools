import os, sys
import fitz

desktop = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
target_form_path = os.path.join(desktop, "CORP-FORM-21 - P1 31 Aug 2026.pdf")

from check_user_layout import t3, t4, t5

print(f"t3 len: {len(t3)}, t4 len: {len(t4)}, t5 len: {len(t5)}")

for fs in [8.0, 8.25, 8.5, 8.75, 9.0, 9.2]:
    doc_tgt = fitz.open(target_form_path)
    
    # P3
    for w in doc_tgt[2].widgets():
        if w.field_name == 'Text Field49':
            w.field_value = t3
            w.text_fontsize = fs
            w.update()
            
    # P4
    for w in doc_tgt[3].widgets():
        if w.field_name == 'Text Field50':
            w.field_value = t4
            w.text_fontsize = fs
            w.update()
            
    # P5
    for w in doc_tgt[4].widgets():
        if w.field_name == 'Text Field51':
            w.field_value = t5
            w.text_fontsize = fs
            w.update()
            
    out_tmp = f"scratch/test_user_layout_{fs}.pdf"
    doc_tgt.save(out_tmp)
    doc_tgt.close()
    
    doc_chk = fitz.open(out_tmp)
    results = []
    for pidx, fn in [(2, 'Text Field49'), (3, 'Text Field50'), (4, 'Text Field51')]:
        p = doc_chk[pidx]
        w_obj = [w for w in p.widgets() if w.field_name == fn][0]
        words = [w for w in p.get_text('words') if w[1] >= w_obj.rect.y0 - 2 and w[3] <= w_obj.rect.y1 + 5]
        last_y = max([w[3] for w in words]) if words else 0
        rem = w_obj.rect.y1 - last_y
        last_w = words[-1][4] if words else ''
        results.append(f"{fn}: rem={rem:.1f}pt ('{last_w}')")
    print(f"fs={fs:4.2f} -> " + " | ".join(results))
    doc_chk.close()
    if os.path.exists(out_tmp):
        os.remove(out_tmp)

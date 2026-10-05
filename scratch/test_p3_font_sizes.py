import fitz

doc_s = fitz.open(r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis.pdf')
tgt_path = r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\CORP-FORM-21 Laboratory OOS Investigation Form (v11.1).pdf'
vals = {w.field_name: w.field_value for p in doc_s for w in p.widgets()}

for fs in [9.0, 9.2, 9.5, 9.8, 10.0]:
    doc_t = fitz.open(tgt_path)
    p3 = doc_t[2]
    for w in p3.widgets():
        if w.field_name in vals and vals[w.field_name]:
            w.field_value = vals[w.field_name]
            if w.field_name == 'Text Field49':
                w.text_fontsize = fs
            w.update()
    doc_t.save(f'scratch/test_p3_fs{fs}.pdf')
    d = fitz.open(f'scratch/test_p3_fs{fs}.pdf')
    p3_out = d[2]
    w49 = [w for w in p3_out.widgets() if w.field_name == 'Text Field49'][0]
    words = [w for w in p3_out.get_text('words') if w[1] >= w49.rect.y0 and w[3] <= w49.rect.y1]
    last_y = max([w[3] for w in words]) if words else 0
    rem = w49.rect.y1 - last_y
    last_word = words[-1][4] if words else ''
    print(f'fs={fs}: words={len(words)}/299, rem={rem:.1f}pt, last_word="{last_word}"')

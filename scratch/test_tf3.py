import fitz

tgt_path = r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\CORP-FORM-21 Laboratory OOS Investigation Form (v11.1).pdf'
doc_s = fitz.open(r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis.pdf')
vals = {w.field_name: w.field_value for p in doc_s for w in p.widgets()}

for fs in [0.0, 7.0, 7.5, 8.0, 8.5]:
    doc_t = fitz.open(tgt_path)
    p1 = doc_t[0]
    for w in p1.widgets():
        if w.field_name in vals and vals[w.field_name]:
            w.field_value = vals[w.field_name]
            if w.field_name == 'Text Field3' and fs > 0:
                w.text_fontsize = fs
            w.update()
    doc_t.save(f'scratch/test_p1_tf3_fs{fs}.pdf')
    d = fitz.open(f'scratch/test_p1_tf3_fs{fs}.pdf')
    pix = d[0].get_pixmap(dpi=150, clip=fitz.Rect(100, 250, 300, 400))
    pix.save(f'scratch/test_p1_tf3_fs{fs}.png')
    
    # check words in Text Field3
    w3 = [w for w in d[0].widgets() if w.field_name == 'Text Field3'][0]
    words = [w[4] for w in d[0].get_text('words') if w[1] >= w3.rect.y0 and w[3] <= w3.rect.y1]
    print(f'fs={fs}: {len(words)} words, content: {" ".join(words)}')

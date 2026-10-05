import fitz

src_path = r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis.pdf'
tgt_path = r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\CORP-FORM-21 Laboratory OOS Investigation Form (v11.1).pdf'

doc_s = fitz.open(src_path)

vals = {}
for p in doc_s:
    for w in p.widgets():
        vals[w.field_name] = w.field_value

for fs in [8.0, 8.2, 8.5, 8.7, 8.8, 9.0, 9.2]:
    doc_t = fitz.open(tgt_path)
    for p in doc_t:
        for w in p.widgets():
            if w.field_name in vals and vals[w.field_name]:
                w.field_value = vals[w.field_name]
                if w.field_name in ['Text Field49', 'Text Field50', 'Text Field51']:
                    w.text_fontsize = fs
                w.update()
    tag = f'corp_test_fs{fs}'
    out_pdf = f'scratch/{tag}.pdf'
    doc_t.save(out_pdf)
    doc_t.close()
    
    d = fitz.open(out_pdf)
    p4 = d[3]
    f50_rect = [w.rect for w in p4.widgets() if w.field_name == 'Text Field50'][0]
    tb4 = p4.get_text('blocks')
    inside4 = [b for b in tb4 if b[1] >= f50_rect.y0 - 5 and b[3] <= f50_rect.y1 + 10]
    last_y4 = max([b[3] for b in inside4]) if inside4 else 0
    last_text4 = inside4[-1][4].strip().split('\n')[-1] if inside4 else ''
    
    p5 = d[4]
    f51_rect = [w.rect for w in p5.widgets() if w.field_name == 'Text Field51'][0]
    tb5 = p5.get_text('blocks')
    inside5 = [b for b in tb5 if b[1] >= f51_rect.y0 - 5 and b[3] <= f51_rect.y1 + 10]
    last_y5 = max([b[3] for b in inside5]) if inside5 else 0
    last_text5 = inside5[-1][4].strip().split('\n')[-1] if inside5 else ''
    
    rem4 = f50_rect.y1 - last_y4
    rem5 = f51_rect.y1 - last_y5
    print(f'fs={fs}: P4 rem={rem4:.1f}pt (last: "{last_text4[-30:]}"), P5 rem={rem5:.1f}pt (last: "{last_text5[-30:]}")')

import fitz

with open('scratch/process_oos_262080.py', 'r', encoding='utf-8') as f:
    code = f.read()

narrative_code = code.split('# 1. NARRATIVE TEXT')[1].split('# ==========================================\n# 2. RENDER WORD DOCUMENTS')[0]
local_scope = {}
exec(narrative_code, local_scope)

p1 = local_scope['p1']
p2 = local_scope['p2']
p3 = local_scope['p3']
p4 = local_scope['p4']
p5 = local_scope['p5']
p5b = local_scope['p5b']
p6 = local_scope['p6']
p7 = local_scope['p7']
p8 = local_scope['p8']
p9 = local_scope['p9']
p10 = local_scope['p10']
p11 = local_scope['p11']
p12 = local_scope['p12']
p13 = local_scope['p13']
p14 = local_scope['p14']
p15 = local_scope['p15']
p16 = local_scope['p16']
p17 = local_scope['p17']
p18 = local_scope['p18']
p19 = local_scope['p19']
p20 = local_scope['p20']
p21 = local_scope['p21']

# splitA: P6-P12 on Page 4 (4578 chars), P13-P21 on Page 5 (4326 chars)
# splitB: P6-P13 on Page 4 (5489 chars), P14-P21 on Page 5 (3415 chars)

splits = [
    ('splitA', [p6, p7, p8, p9, p10, p11, p12], [p13, p14, p15, p16, p17, p18, p19, p20, p21]),
]

for name, p4_list, p5_list in splits:
    t4 = '\n\n'.join(p4_list)
    t5 = '\n\n'.join(p5_list)
    print(f'{name}: P4 len={len(t4)}, P5 len={len(t5)}')
    for fs in [8.5, 9.0, 9.5]:
        doc = fitz.open(r'Celsis OOS P1 template.pdf')
        
        page4 = doc[3]
        for w in page4.widgets():
            if w.field_name == 'Text Field50':
                w.field_value = t4.replace('\n\n', '\r \r').replace('\n', '\r')
                w.text_fontsize = fs
                w.update()
                
        page5 = doc[4]
        for w in page5.widgets():
            if w.field_name == 'Text Field51':
                w.field_value = t5.replace('\n\n', '\r \r').replace('\n', '\r')
                w.text_fontsize = fs
                w.update()
                
        out_pdf = f'scratch/{name}_fs{fs}.pdf'
        doc.save(out_pdf)
        doc.close()
        
        d = fitz.open(out_pdf)
        d[3].get_pixmap(dpi=150).save(f'scratch/{name}_fs{fs}_p4.png')
        d[4].get_pixmap(dpi=150).save(f'scratch/{name}_fs{fs}_p5.png')
        d.close()
        print(f'Saved {name} at fs={fs}')

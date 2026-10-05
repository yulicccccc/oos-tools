import fitz

p8 = (
    "Following the reading, sample ETX-260828-0527 was found to yield a positive reading in one of the 300 mL FTM media jars "
    "(ETX-260828-0527-3/5). The average Relative Luminescence Units (RLU) from the duplicate reading tube, originating from "
    "the FTM sample bottle, yielded 7,190 RLU, which exceeded the method cutoff of 2,955 RLU, where the FTM negative control "
    "was 985 RLU. All other FTM and TSB sample bottles in the same batch tested negative. The %CV from the duplicate reading tubes "
    "for the positive FTM bottles was well within the specification (< 30%). Additionally, all Daily Controls, including the "
    "Instrument Blank, Reagent Blank, and ATP Positive Control, were within the defined specifications, each with a %CV below 30%."
)

p19 = 'Analyzing a 6-month sample history for Optimal Balance Pharmacy indicates, this specific analyte "MOTs-C 10 MG/ML (5 ML) Injection" has had no prior failures using the Celsis Sterility testing during this period.'

with open('scratch/process_oos_262080.py', 'r', encoding='utf-8') as f:
    code = f.read()

local_scope = {}
exec(code.split('# 1. NARRATIVE TEXT')[1].split('# ==========================================\n# 2. RENDER WORD DOCUMENTS')[0], local_scope)

local_scope['p8'] = p8
local_scope['p19'] = p19

p4_list = [local_scope[f'p{i}'] for i in [6, 7, 8, 9, 10, 11, 12]]
p5_list = [local_scope[f'p{i}'] for i in [13, 14, 15, 16, 17, 18, 19, 21]]

t4 = '\n\n'.join(p4_list)
t5 = '\n\n'.join(p5_list)

print('P4 len:', len(t4))
print('P5 len:', len(t5))

for fs in [10.3, 10.4, 10.5, 10.6, 10.8]:
    doc = fitz.open('Celsis OOS P1 template.pdf')
    page4 = doc[3]
    for w in page4.widgets():
        if w.field_name == 'Text Field50':
            w.field_value = t4.replace('₂', '2').replace('\n\n', '\r \r').replace('\n', '\r')
            w.text_fontsize = fs
            w.update()
    page5 = doc[4]
    for w in page5.widgets():
        if w.field_name == 'Text Field51':
            w.field_value = t5.replace('\n\n', '\r \r').replace('\n', '\r')
            w.text_fontsize = fs
            w.update()
    tag = f'test_3ch_fs{fs}'
    out_pdf = f'scratch/{tag}.pdf'
    doc.save(out_pdf)
    doc.close()
    
    d = fitz.open(out_pdf)
    f50_rect = [w.rect for w in d[3].widgets() if w.field_name == 'Text Field50'][0]
    tb4 = d[3].get_text('blocks')
    inside4 = [b for b in tb4 if b[1] >= f50_rect.y0 - 5 and b[3] <= f50_rect.y1 + 10]
    last_y4 = max([b[3] for b in inside4]) if inside4 else 0
    last_text4 = inside4[-1][4].strip().split('\n')[-1] if inside4 else ''
    
    f51_rect = [w.rect for w in d[4].widgets() if w.field_name == 'Text Field51'][0]
    tb5 = d[4].get_text('blocks')
    inside5 = [b for b in tb5 if b[1] >= f51_rect.y0 - 5 and b[3] <= f51_rect.y1 + 10]
    last_y5 = max([b[3] for b in inside5]) if inside5 else 0
    last_text5 = inside5[-1][4].strip().split('\n')[-1] if inside5 else ''
    
    d[3].get_pixmap(dpi=150).save(f'scratch/{tag}_p4.png')
    d[4].get_pixmap(dpi=150).save(f'scratch/{tag}_p5.png')
    d.close()
    
    rem4 = f50_rect.y1 - last_y4
    rem5 = f51_rect.y1 - last_y5
    print(f'fs={fs}: P4 rem={rem4:.1f}pt (last: "{last_text4[-30:]}"), P5 rem={rem5:.1f}pt (last: "{last_text5[-30:]}")')

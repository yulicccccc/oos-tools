import fitz

for fs in [10.2, 10.3, 10.4, 10.5]:
    tag = f'opt1_fs{fs}'
    doc = fitz.open(f'scratch/{tag}.pdf')
    
    p4 = doc[3]
    f50_rect = None
    for w in p4.widgets():
        if w.field_name == 'Text Field50':
            f50_rect = w.rect
            
    p5 = doc[4]
    f51_rect = None
    for w in p5.widgets():
        if w.field_name == 'Text Field51':
            f51_rect = w.rect
            
    tb4 = p4.get_text('blocks')
    inside4 = [b for b in tb4 if b[1] >= f50_rect.y0 - 5 and b[3] <= f50_rect.y1 + 10]
    last_y4 = max([b[3] for b in inside4]) if inside4 else 0
    last_text4 = inside4[-1][4].strip().split('\n')[-1] if inside4 else ''
    
    tb5 = p5.get_text('blocks')
    inside5 = [b for b in tb5 if b[1] >= f51_rect.y0 - 5 and b[3] <= f51_rect.y1 + 10]
    last_y5 = max([b[3] for b in inside5]) if inside5 else 0
    last_text5 = inside5[-1][4].strip().split('\n')[-1] if inside5 else ''
    
    rem4 = f50_rect.y1 - last_y4
    rem5 = f51_rect.y1 - last_y5
    print(f'fs={fs}: P4 rem={rem4:.1f}pt (last: "{last_text4[-30:]}"), P5 rem={rem5:.1f}pt (last: "{last_text5[-30:]}")')

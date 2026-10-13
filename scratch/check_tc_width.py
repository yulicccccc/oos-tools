import docx

doc = docx.Document('scratch/Celsis table OOS-262017.docx')
t2 = doc.tables[1]

for r_name, r in [('R2', t2.rows[2]), ('R3', t2.rows[3])]:
    print(f"=== {r_name} ===")
    for i, c in enumerate(r._tr.xpath('w:tc')):
        txt = ''.join(c.xpath('.//w:t/text()'))
        span = c.xpath('.//w:gridSpan/@w:val')
        w = c.xpath('.//w:tcW/@w:w')
        jc = c.xpath('.//w:jc/@w:val')
        print(f"  tc {i}: w={w}, span={span}, jc={jc}, txt='{txt}'")

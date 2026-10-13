import docx

doc = docx.Document('scratch/Celsis table OOS-262017.docx')
t2 = doc.tables[1]

print("=== R2 ===")
for i, c in enumerate(t2.rows[2]._tr.xpath('w:tc')):
    txt = ''.join(c.xpath('.//w:t/text()'))
    span = c.xpath('.//w:gridSpan/@w:val')
    print(f"  tc {i}: text='{txt}', gridSpan={span}")

print("=== R3 ===")
for i, c in enumerate(t2.rows[3]._tr.xpath('w:tc')):
    txt = ''.join(c.xpath('.//w:t/text()'))
    span = c.xpath('.//w:gridSpan/@w:val')
    print(f"  tc {i}: text='{txt}', gridSpan={span}")

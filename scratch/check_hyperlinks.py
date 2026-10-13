import docx
from docx.oxml.ns import qn

doc = docx.Document(r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\Tables OOS-262237 GoGoMeds Select (E10747) - ScanRDI.docx')
print('=== Relationships ===')
for r_id, rel in doc.part.rels.items():
    if 'hyperlink' in rel.reltype:
        print(f'{r_id}: target={rel.target_ref}')

print('\n=== Hyperlink tags in XML ===')
for t_idx, table in enumerate(doc.tables):
    for r_idx, row in enumerate(table.rows):
        for c_idx, cell in enumerate(row.cells):
            hyperlinks = cell._element.xpath('.//w:hyperlink')
            if hyperlinks:
                for h in hyperlinks:
                    r_id = h.get(qn('r:id'))
                    h_text = ''.join(h.itertext())
                    target = doc.part.rels[r_id].target_ref if r_id in doc.part.rels else 'UNKNOWN'
                    print(f'T{t_idx+1} R{r_idx} C{c_idx}: text="{h_text}" r:id={r_id} target={target}')

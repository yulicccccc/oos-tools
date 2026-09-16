import json, docx
from docxtpl import DocxTemplate

with open('scratch/UPDATED_SAVE_OOS-261814 GoGoMeds Select (E10747) - ScanRDI.txt', 'r', encoding='utf-8') as f:
    data = json.load(f)

data['note_pers'] = 'plate was desiccated on 5 day read.'
data['note_surf'] = 'None'
data['note_sett'] = 'plate was desiccated on 5 day read.'
data['note_air'] = 'None'
data['note_room'] = 'None'
data['note_surf_chg'] = 'None'
data['note_sett_chg'] = 'plate was desiccated on 5 day read.'

bsc_id = str(data.get('bsc_id','')).strip()
chgbsc_id = str(data.get('chgbsc_id','')).strip()
test_date = data.get('test_date', '')

if bsc_id == chgbsc_id or not chgbsc_id or chgbsc_id == 'N/A':
    data['smart_bsc_bracketing_header'] = f'Biological Safety Cabinet EM Bracketing Biological Safety Cabinet (BSC) E00{bsc_id} on {test_date}'
else:
    data['smart_bsc_bracketing_header'] = f'Biological Safety Cabinet EM Bracketing Biological Safety Cabinet (BSC) E00{bsc_id} and E00{chgbsc_id} on {test_date}'

tpl = DocxTemplate('scratch/test_tables_for_scan_updated.docx')
tpl.render(data)

if bsc_id == chgbsc_id or not chgbsc_id or chgbsc_id == 'N/A':
    t = tpl.docx.tables[1]
    t._tbl.remove(t.rows[7]._tr)
    t._tbl.remove(t.rows[5]._tr)

tpl.save('scratch/test_rendered_table_final.docx')

doc = docx.Document('scratch/test_rendered_table_final.docx')
t = doc.tables[1]
print(f'Final Table 2 Rows: {len(t.rows)}')
for i, r in enumerate(t.rows):
    print(f'Row {i:2d}: site={r.cells[0].text.strip()[:35]:35s} | obs={r.cells[8].text.strip():20s} | notes={r.cells[-1].text.strip()}')

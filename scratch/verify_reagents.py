import pypdf
import docx

# PDF check
r = pypdf.PdfReader('scratch/OOS_261186_Report.pdf')
f = r.get_fields()
print('PDF Text Field22 (Lot):', repr(f.get('Text Field22', {}).get('/V')))
print('PDF Text Field23 (Exp):', repr(f.get('Text Field23', {}).get('/V')))

# Word check
doc = docx.Document('scratch/OOS_261186_Report.docx')
t0 = doc.tables[0]
r23 = t0.rows[23]
print('Word Row 23 Cell 6 (Lot):', repr(r23.cells[6].text))
print('Word Row 23 Cell 11 (Exp):', repr(r23.cells[11].text))

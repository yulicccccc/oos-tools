import pypdf
import docx

# PDF check
r = pypdf.PdfReader('scratch/OOS_261186_Report.pdf')
f = r.get_fields()
print('PDF Cleanroom Room/Eq (Text Field32):', repr(f.get('Text Field32', {}).get('/V')))
print('PDF Cleanroom Exp (Text Field33):', repr(f.get('Text Field33', {}).get('/V')))
print('PDF System Check Room/Eq (Text Field34):', repr(f.get('Text Field34', {}).get('/V')))
print('PDF System Check Exp (Text Field35):', repr(f.get('Text Field35', {}).get('/V')))

# Word check
doc = docx.Document('scratch/OOS_261186_Report.docx')
t0 = doc.tables[0]
r29 = t0.rows[29]
r30 = t0.rows[30]
r31 = t0.rows[31]
print('Word Row 29 Cell 6 / 11:', repr(r29.cells[6].text), '/', repr(r29.cells[11].text))
print('Word Row 30 Cell 6 / 11:', repr(r30.cells[6].text), '/', repr(r30.cells[11].text))
print('Word Row 31 Cell 6 / 11:', repr(r31.cells[6].text), '/', repr(r31.cells[11].text))

import pypdf
import docx

reader = pypdf.PdfReader('scratch/OOS_261186_Report.pdf')
fields = reader.get_fields()
print('PDF Check Box0 (Client Care):', repr(fields.get('Check Box0', {}).get('/V')))
print('PDF Check Box1 (Client 24h):', repr(fields.get('Check Box1', {}).get('/V')))

doc = docx.Document('scratch/OOS_261186_Report.docx')
t0 = doc.tables[0]
r5 = t0.rows[5]
print('Word Row 5 has default=0:', 'w:val="0"' in r5.cells[5]._tc.xml)
print('Word Row 5 has default=1:', 'w:val="1"' in r5.cells[5]._tc.xml)

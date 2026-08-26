import pypdf
import docx
import re

# PDF check
r = pypdf.PdfReader('scratch/OOS_261186_Report.pdf')
f = r.get_fields()
print('PDF Integrity (Yes, No, N/A):', [f.get(f'Check Box{i}', {}).get('/V') for i in [31, 32, 33]])
print('PDF Transport (Yes, No, N/A):', [f.get(f'Check Box{i}', {}).get('/V') for i in [34, 35, 36]])
print('PDF Run Results (Yes, No, N/A):', [f.get(f'Check Box{i}', {}).get('/V') for i in [37, 38, 39]])

# Word check
doc = docx.Document('scratch/OOS_261186_Report.docx')
t0 = doc.tables[0]
r20_defaults = re.findall(r'<w:default w:val="(\d+)"/>', t0.rows[20]._tr.xml)
r21_defaults = re.findall(r'<w:default w:val="(\d+)"/>', t0.rows[21]._tr.xml)
print('Word Row 20 defaults (Integrity Yes, No, N/A):', r20_defaults)
print('Word Row 21 defaults (Run results Yes, No, N/A):', r21_defaults)

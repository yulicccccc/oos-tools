import pypdf
import docx
import re

# PDF check
r = pypdf.PdfReader('scratch/OOS_261186_Report.pdf')
f = r.get_fields()
print('PDF Check Box7 (Yes):', repr(f.get('Check Box7', {}).get('/V')))
print('PDF Check Box8 (No):', repr(f.get('Check Box8', {}).get('/V')))
print('PDF Check Box9 (N/A):', repr(f.get('Check Box9', {}).get('/V')))

# Word check
doc = docx.Document('scratch/OOS_261186_Report.docx')
t0 = doc.tables[0]
r11 = t0.rows[11]
defaults = re.findall(r'<w:default w:val="(\d+)"/>', r11._tr.xml)
print('Word Row 11 defaults (Yes, No, N/A):', defaults)

import os
from pypdf import PdfReader

src_pdf = r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\OOS-261187 ScanC_O CGS E001309 S1 11MAY2026 - EM.pdf'
r = PdfReader(src_pdf)
fields = r.get_fields()
print('Total fields:', len(fields))

t_fields = {k: v.get('/V') for k, v in fields.items() if k.startswith('Text Field')}
d_fields = {k: v.get('/V') for k, v in fields.items() if k.startswith('Date Field')}
c_fields = {k: v.get('/V') for k, v in fields.items() if k.startswith('Check Box')}

print(f"Text fields count: {len(t_fields)}, non-empty: {sum(1 for v in t_fields.values() if v not in [None, ''])}")
print(f"Date fields count: {len(d_fields)}, non-empty: {sum(1 for v in d_fields.values() if v not in [None, ''])}")
print(f"Checkbox fields count: {len(c_fields)}, checked: {sum(1 for v in c_fields.values() if v in ['/Yes', '/1', True])}")

print("\n--- All Non-Empty Fields ---")
for k, v in sorted(fields.items()):
    val = v.get('/V')
    if val not in [None, '', '/Off']:
        val_str = str(val).replace('\n', ' \\n ')
        if len(val_str) > 70:
            val_str = val_str[:70] + '...'
        print(f"{k:20s}: {val_str}")

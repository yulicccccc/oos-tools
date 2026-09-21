import os
from pypdf import PdfReader

path = r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\CORP-FORM-21 Laboratory OOS Investigation Form (v11.1).pdf'
reader = PdfReader(path)
print('Num pages:', len(reader.pages))
fields = reader.get_fields()
print('Num fields:', len(fields) if fields else 0)
if fields:
    for i, (k, v) in enumerate(fields.items()):
        ft = v.get('/FT', '')
        val = v.get('/V', '')
        if ft == '/Sig':
            val_str = '<Signature>'
        else:
            val_str = repr(val)
            if len(val_str) > 60:
                val_str = val_str[:60] + '...'
        print(f"[{i:3d}] Field: {k!r:35s} | Type: {ft!r:10s} | Val: {val_str}")

import os
from pypdf import PdfReader

path = r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\CORP-FORM-21 Laboratory OOS Investigation Form (v11.1).pdf'
print('Exists:', os.path.exists(path))
if os.path.exists(path):
    reader = PdfReader(path)
    print('Num pages:', len(reader.pages))
    fields = reader.get_fields()
    print('Num fields:', len(fields) if fields else 0)
    if fields:
        for i, (k, v) in enumerate(fields.items()):
            val = v.get('/V', '')
            ft = v.get('/FT', '')
            print(f"[{i}] {k}: val='{val}' ft='{ft}'")
    else:
        print("No interactive form fields found! Let's check text on pages:")
        for idx, page in enumerate(reader.pages):
            text = page.extract_text() or ''
            print(f"--- Page {idx+1} ({len(text)} chars) ---")
            print(text[:300].replace('\n', ' '))

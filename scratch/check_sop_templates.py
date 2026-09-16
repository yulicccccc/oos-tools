import os, glob, docx
from pypdf import PdfReader

print('=== DOCX FILES ===')
for f in glob.glob('ScanRDI*.docx'):
    try:
        d = docx.Document(f)
        matches = set()
        for p in d.paragraphs:
            if '2.600.023' in p.text or '2.700.004' in p.text:
                matches.add(p.text.strip())
        for t in d.tables:
            for row in t.rows:
                for c in row.cells:
                    if '2.600.023' in c.text or '2.700.004' in c.text:
                        matches.add(c.text.strip().replace('\n', ' '))
        if matches:
            print(f'*** {f} ***')
            for m in matches:
                print('  -', m[:100])
        else:
            print(f'{f}: None')
    except Exception as e:
        print(f'{f}: error {e}')

print('\n=== PDF FILES ===')
for f in glob.glob('ScanRDI*.pdf'):
    try:
        r = PdfReader(f)
        fields = r.get_fields() or {}
        matches = []
        for k, v in fields.items():
            val = str(v.get('/V', ''))
            if '2.600.023' in val or '2.700.004' in val:
                matches.append((k, val))
        if matches:
            print(f'*** {f} ***')
            for k, val in matches:
                print(f'  {k}: {repr(val)}')
        else:
            print(f'{f}: None')
    except Exception as e:
        print(f'{f}: error {e}')

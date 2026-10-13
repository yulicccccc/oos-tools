import os
import sys
import shutil
from datetime import datetime
import fitz

desktop = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
target_form_path = os.path.join(desktop, "CORP-FORM-21 - P1 31 Aug 2026.pdf")
source_pdf_path = os.path.join(desktop, "OOS-262017 Revive Rx Pharmacy - PO Required (E00927) - Celsis.pdf")
tables_pdf_path = os.path.join(desktop, "Celsis table OOS-262017.pdf")

doc_src = fitz.open(source_pdf_path)
doc_tgt = fitz.open(target_form_path)

src_fields = {}
for p in doc_src:
    for w in p.widgets():
        src_fields[w.field_name] = (w.field_value, w.text_fontsize, w.field_type_string)

print(f"Total source fields: {len(src_fields)}")

# Now populate target
mismatches = 0
for pno in range(len(doc_tgt)):
    p = doc_tgt[pno]
    for w in p.widgets():
        fn = w.field_name
        if fn in src_fields:
            val, fs, ftype = src_fields[fn]
            if ftype == 'CheckBox':
                w.field_value = val if val in ['Yes', 'Off'] else ('Yes' if val else 'Off')
            else:
                w.field_value = val
                if fs > 0:
                    w.text_fontsize = fs
            w.update()

out_test = r"scratch\test_transcribed_corp.pdf"
doc_tgt.save(out_test)
doc_tgt.close()

# Verify
doc_ver = fitz.open(out_test)
for pno in range(len(doc_ver)):
    for w in doc_ver[pno].widgets():
        fn = w.field_name
        if fn in src_fields:
            s_val, s_fs, s_type = src_fields[fn]
            c_val = w.field_value
            if s_val != c_val:
                print(f"Mismatch in {fn}: src={repr(s_val)} vs cur={repr(c_val)}")
                mismatches += 1

print(f"Verification complete: {mismatches} mismatches.")

# Check text bounds on Pages 3, 4, 5
for page_idx, fn in [(2, 'Text Field49'), (3, 'Text Field50'), (4, 'Text Field51')]:
    p = doc_ver[page_idx]
    w_obj = [w for w in p.widgets() if w.field_name == fn][0]
    words = [w for w in p.get_text('words') if w[1] >= w_obj.rect.y0 - 2 and w[3] <= w_obj.rect.y1 + 5]
    last_y = max([w[3] for w in words]) if words else 0
    rem = w_obj.rect.y1 - last_y
    last_word = words[-1][4] if words else ''
    print(f"  Page {page_idx+1} ({fn}): {len(words)} words, bottom rem={rem:.1f}pt, last word='{last_word}'")
    assert rem > 0, f"Text overflow on Page {page_idx+1} ({fn})!"

# Check Page 3 Text Field48
p3_w48 = [w for w in doc_ver[2].widgets() if w.field_name == 'Text Field48'][0]
print(f"Page 3 Text Field48 value: '{p3_w48.field_value}'")

doc_ver.close()

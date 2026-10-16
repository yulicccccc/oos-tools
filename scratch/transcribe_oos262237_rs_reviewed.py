import os
import sys
import shutil
from datetime import datetime
import fitz

sys.stdout.reconfigure(encoding='utf-8')

DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
DOCS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
HISTORY_DIR = os.path.join(DOCS_DIR, ".history")
SCRATCH_DIR = os.path.join(DOCS_DIR, "scratch")

os.makedirs(HISTORY_DIR, exist_ok=True)
os.makedirs(SCRATCH_DIR, exist_ok=True)

target_corp_path = os.path.join(DESKTOP_DIR, "CORP-FORM-21 Laboratory OOS Investigation Form (v11.1) (3).pdf")
source_oos_pdf = os.path.join(DESKTOP_DIR, "OOS-262237 GoGoMeds Select (E10747) - ScanRDI  -RS Reviewed -  Signed by.pdf")

print("=== STEP 1: BACKUP ORIGINAL CORP-FORM-21 (3) ===")
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
backup_name = f"CORP-FORM-21 Laboratory OOS Investigation Form (v11.1) (3)_backup_{timestamp}.pdf"
backup_path = os.path.join(HISTORY_DIR, backup_name)
shutil.copy2(target_corp_path, backup_path)
print(f"Backed up pristine template to: {backup_path}")

print("\n=== STEP 2: EXTRACT FIELDS AND STYLING FROM RS REVIEWED SOURCE ===")
doc_src = fitz.open(source_oos_pdf)
print(f"Source PDF total pages: {len(doc_src)}")

src_fields = {}
for p in doc_src:
    for w in p.widgets():
        src_fields[w.field_name] = (w.field_value, w.text_fontsize, w.field_type_string)

print(f"Extracted {len(src_fields)} fields from RS Reviewed source PDF.")

print("\n=== STEP 3: TRANSCRIBE INTO FRESH CORP-FORM-21 ===")
doc_tgt = fitz.open(backup_path)
assert len(doc_tgt) == 6, f"Expected 6 pages in base CORP-FORM-21, got {len(doc_tgt)}"

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
                if fn == 'Text Field49':
                    w.text_fontsize = 9.1  # Calibrate to ensure comfortable bottom margin
                elif fs > 0:
                    w.text_fontsize = fs
            w.update()

# Step 4: Append Page 7 (Table 1 & Table 2) with active links
if len(doc_src) >= 7:
    doc_tgt.insert_pdf(doc_src, from_page=6, to_page=6)
    print("Appended Table 1 & Table 2 as Page 7.")

# Save to scratch first
scratch_transcribed = os.path.join(SCRATCH_DIR, "CORP-FORM-21_262237_transcribed.pdf")
doc_tgt.save(scratch_transcribed)
doc_tgt.close()
print(f"Saved transcribed form to: {scratch_transcribed}")

print("\n=== STEP 4: VERIFY TRANSCRIBED PDF ===")
doc_ver = fitz.open(scratch_transcribed)
print(f"Total pages in transcribed PDF: {len(doc_ver)}")
assert len(doc_ver) == 7, f"Expected exactly 7 pages, got {len(doc_ver)}"

ver_fields = {w.field_name: w.field_value for p in doc_ver for w in p.widgets()}
mismatches = 0
for fn, (s_val, s_fs, s_ft) in src_fields.items():
    c_val = ver_fields.get(fn)
    # Checkbox equality normalize
    if s_ft == 'CheckBox':
        norm_s = 'Yes' if s_val in ['Yes', True] else 'Off'
        norm_c = 'Yes' if c_val in ['Yes', True] else 'Off'
        if norm_s != norm_c:
            print(f"  Checkbox mismatch in {fn}: src='{s_val}' != cur='{c_val}'")
            mismatches += 1
    else:
        if s_val != c_val:
            print(f"  Text mismatch in {fn}: src='{repr(s_val)}' != cur='{repr(c_val)}'")
            mismatches += 1

print(f"Field value verification complete: {mismatches} mismatches.")
assert mismatches == 0, f"Found {mismatches} field mismatches!"

# Verify hyperlinks on Page 7
p7_links = doc_ver[6].get_links()
print(f"Page 7 verified active hyperlinks count: {len(p7_links)}")
for l in p7_links:
    print(f"  Link: {l.get('uri')}")
assert len(p7_links) >= 5, f"Expected at least 5 links on Page 7, got {len(p7_links)}"

# Check text bounds on Pages 3, 4, 5
for page_idx, fn in [(2, 'Text Field49'), (3, 'Text Field50'), (4, 'Text Field51')]:
    p = doc_ver[page_idx]
    w_obj = [w for w in p.widgets() if w.field_name == fn][0]
    words = [w for w in p.get_text('words') if w[1] >= w_obj.rect.y0 - 2 and w[3] <= w_obj.rect.y1 + 5]
    last_y = max([w[3] for w in words]) if words else 0
    rem = w_obj.rect.y1 - last_y
    last_word = words[-1][4] if words else ''
    print(f"  Page {page_idx+1} ({fn}): {len(words)} words, bottom remaining={rem:.1f}pt, last word='{last_word}'")
    assert rem > 0, f"Text overflow on Page {page_idx+1} ({fn})!"

doc_ver.close()

print("\n=== STEP 5: SYNC TO DESKTOP AND VERSION CONTROL ===")
# 1. Update Desktop target file
shutil.copy2(scratch_transcribed, target_corp_path)
print(f"Successfully updated Desktop file: {target_corp_path}")

# 2. Also save descriptive master name to Desktop
descriptive_desktop = os.path.join(DESKTOP_DIR, "CORP-FORM-21 - OOS-262237 GoGoMeds Select (E10747) - Phase I.pdf")
shutil.copy2(scratch_transcribed, descriptive_desktop)
print(f"Saved descriptive copy to Desktop: {descriptive_desktop}")

# 3. Save copy in Documents\OOS for git tracking
docs_archive = os.path.join(DOCS_DIR, "CORP-FORM-21 - OOS-262237 GoGoMeds Select (E10747) - Phase I.pdf")
shutil.copy2(scratch_transcribed, docs_archive)
print(f"Saved version-controlled copy to Documents: {docs_archive}")

print("\n=== STEP 6: RENDER VERIFICATION IMAGES ===")
doc_final = fitz.open(scratch_transcribed)
for idx, page in enumerate(doc_final):
    pix = page.get_pixmap(dpi=150)
    img_out = os.path.join(SCRATCH_DIR, f"transcribed_262237_p{idx+1}.png")
    pix.save(img_out)
    print(f"Rendered Page {idx+1} to {img_out}")

doc_final.close()
print("\n=== ALL TRANSCRIPTION & VERIFICATION TASKS COMPLETED SUCCESSFULLY! ===")

import os
import sys
import shutil
from datetime import datetime
import fitz  # PyMuPDF

sys.stdout.reconfigure(encoding='utf-8')

DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
DOCS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
HISTORY_DIR = os.path.join(DOCS_DIR, ".history")
SCRATCH_DIR = os.path.join(DOCS_DIR, "scratch")

os.makedirs(HISTORY_DIR, exist_ok=True)
os.makedirs(SCRATCH_DIR, exist_ok=True)

target_corp_path = os.path.join(DESKTOP_DIR, "CORP-FORM-21 Laboratory OOS Investigation Form (v11.1).pdf")
source_oos_pdf = os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis.pdf")
tables_pdf_path = os.path.join(DESKTOP_DIR, "Celsis table OOS-262080.pdf")
complete_pdf_path = os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis (Complete).pdf")

print("=== STEP 1: BACKUP ORIGINAL CORP-FORM-21 ===")
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
backup_name = f"CORP-FORM-21 Laboratory OOS Investigation Form (v11.1)_backup_{timestamp}.pdf"
backup_path = os.path.join(HISTORY_DIR, backup_name)
shutil.copy2(target_corp_path, backup_path)
print(f"Backed up to: {backup_path}")

print("\n=== STEP 2: EXTRACT FIELDS FROM SOURCE OOS PDF ===")
doc_src = fitz.open(source_oos_pdf)
print(f"Source PDF page count: {len(doc_src)}")

field_values = {}
field_types = {}
checkbox_states = {}

for p_no, page in enumerate(doc_src):
    for w in page.widgets():
        field_values[w.field_name] = w.field_value
        field_types[w.field_name] = w.field_type_string
        if w.field_type_string == 'CheckBox':
            checkbox_states[w.field_name] = w.field_value

print(f"Extracted {len(field_values)} fields from source PDF.")
checked_count = sum(1 for v in checkbox_states.values() if v != 'Off')
print(f"Source has {checked_count} checked checkboxes.")

print("\n=== STEP 3: TRANSCRIBE INTO FRESH CORP-FORM-21 ===")
# We open from backup to ensure a pristine base template
doc_tgt = fitz.open(backup_path)
assert len(doc_tgt) == 6, f"Expected 6 pages in base CORP-FORM-21, got {len(doc_tgt)}"

for p_no, page in enumerate(doc_tgt):
    for w in page.widgets():
        fn = w.field_name
        if fn in field_values:
            val = field_values[fn]
            w.field_value = val
            
            # Apply golden calibrated font sizes for multiline narrative fields
            # on CORP-FORM-21 to prevent any bottom text cutoff or void
            if fn in ['Text Field49', 'Text Field50', 'Text Field51']:
                w.text_fontsize = 9.2
            w.update()

scratch_transcribed = os.path.join(SCRATCH_DIR, "CORP-FORM-21_transcribed.pdf")
doc_tgt.save(scratch_transcribed)
doc_tgt.close()
print(f"Saved transcribed form to: {scratch_transcribed}")

print("\n=== STEP 4: VERIFY TRANSCRIBED PDF ===")
doc_ver = fitz.open(scratch_transcribed)
ver_vals = {w.field_name: w.field_value for p in doc_ver for w in p.widgets()}
ver_cb = sum(1 for p in doc_ver for w in p.widgets() if w.field_type_string == 'CheckBox' and w.field_value != 'Off')

mismatches = 0
for fn, orig_val in field_values.items():
    cur_val = ver_vals.get(fn)
    if orig_val != cur_val:
        print(f"  MISMATCH in {fn}: orig='{orig_val}' != cur='{cur_val}'")
        mismatches += 1

print(f"Verification: {len(ver_vals)} fields total, {mismatches} mismatches.")
print(f"Verification: {ver_cb} checked checkboxes (expected {checked_count}).")
assert mismatches == 0, f"Found {mismatches} field mismatches!"
assert ver_cb == checked_count, f"Checkbox mismatch: {ver_cb} vs {checked_count}"

# Check remaining space and bounds on Pages 3, 4, 5
for page_idx, fn in [(2, 'Text Field49'), (3, 'Text Field50'), (4, 'Text Field51')]:
    p = doc_ver[page_idx]
    w_obj = [w for w in p.widgets() if w.field_name == fn][0]
    words = [w for w in p.get_text('words') if w[1] >= w_obj.rect.y0 - 2 and w[3] <= w_obj.rect.y1 + 5]
    last_y = max([w[3] for w in words]) if words else 0
    rem = w_obj.rect.y1 - last_y
    last_word = words[-1][4] if words else ''
    print(f"  Page {page_idx+1} ({fn}): {len(words)} words, bottom rem={rem:.1f}pt, last word='{last_word}'")
    assert rem > 0, f"Text overflow on Page {page_idx+1} ({fn})!"

doc_ver.close()

print("\n=== STEP 5: ASSEMBLE COMBINED 8-PAGE PDF ===")
doc_combined = fitz.open()
doc_trans = fitz.open(scratch_transcribed)
doc_combined.insert_pdf(doc_trans)
doc_trans.close()

if os.path.exists(tables_pdf_path):
    doc_tables = fitz.open(tables_pdf_path)
    doc_combined.insert_pdf(doc_tables)
    doc_tables.close()
    print(f"Appended tables from {tables_pdf_path}. Total pages: {len(doc_combined)}")
else:
    print(f"Warning: {tables_pdf_path} not found!")

scratch_complete = os.path.join(SCRATCH_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis (Complete).pdf")
doc_combined.save(scratch_complete)
doc_combined.close()
print(f"Saved complete combined PDF to: {scratch_complete}")

print("\n=== STEP 6: SYNC DELIVERABLES TO DESKTOP AND DOCUMENTS ===")
def safe_copy(src, dst):
    try:
        shutil.copy2(src, dst)
        print(f"Successfully copied to: {dst}")
        return True
    except PermissionError:
        print(f"WARNING: File lock on {dst}. File is open in another application.")
        return False

# 1. Update CORP-FORM-21 on Desktop
safe_copy(scratch_transcribed, target_corp_path)

# Also copy to Documents\OOS for version control
docs_corp_path = os.path.join(DOCS_DIR, "CORP-FORM-21 Laboratory OOS Investigation Form (v11.1).pdf")
safe_copy(scratch_transcribed, docs_corp_path)

# 2. Update Complete PDF on Desktop and Documents
safe_copy(scratch_complete, complete_pdf_path)
docs_complete_path = os.path.join(DOCS_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis (Complete).pdf")
safe_copy(scratch_complete, docs_complete_path)

print("\n=== STEP 7: RENDER FINAL VERIFICATION IMAGES ===")
doc_final = fitz.open(scratch_complete if os.path.exists(scratch_complete) else scratch_transcribed)
for idx, page in enumerate(doc_final):
    pix = page.get_pixmap(dpi=150)
    img_out = os.path.join(SCRATCH_DIR, f"final_transcribed_p{idx+1}.png")
    pix.save(img_out)
    print(f"Rendered Page {idx+1} to {img_out}")

print("\n=== TRANSCRIPTION AND PACKAGING COMPLETE! ===")

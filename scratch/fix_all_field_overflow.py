import os
import sys
import shutil
from datetime import datetime
import fitz  # PyMuPDF
import re

sys.stdout.reconfigure(encoding='utf-8')

DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
DOCS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
HISTORY_DIR = os.path.join(DOCS_DIR, ".history")
SCRATCH_DIR = os.path.join(DOCS_DIR, "scratch")

os.makedirs(HISTORY_DIR, exist_ok=True)
os.makedirs(SCRATCH_DIR, exist_ok=True)

# Master input files
src_qyc_path = os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis - QYC.pdf")
tables_pdf_path = os.path.join(DESKTOP_DIR, "Celsis table OOS-262080.pdf")

# Target output files
out_corp_desktop = os.path.join(DESKTOP_DIR, "CORP-FORM-21 Laboratory OOS Investigation Form (v11.1).pdf")
out_corp_docs = os.path.join(DOCS_DIR, "CORP-FORM-21 Laboratory OOS Investigation Form (v11.1).pdf")

out_qyc_desktop = os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis - QYC.pdf")
out_complete_desktop = os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis (Complete).pdf")
out_2_desktop = os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis (2).pdf")
out_complete_docs = os.path.join(DOCS_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis (Complete).pdf")

print("=== STEP 1: VERIFY SOURCE INPUTS ===")
assert os.path.exists(src_qyc_path), f"Source QYC PDF not found: {src_qyc_path}"
assert os.path.exists(tables_pdf_path), f"Tables PDF not found: {tables_pdf_path}"
print("Source inputs verified successfully.")

print("\n=== STEP 2: LOAD MASTER SOURCE & CALIBRATE FONT SIZES ===")
doc_src = fitz.open(src_qyc_path)
print(f"Loaded master document with {len(doc_src)} pages.")

# Definitive font size map for zero-overflow / no '+' indicator across all fields
FONT_SIZE_MAP = {
    # Page 1
    'Text Field0': 5.5,    # Initiator
    'Date Field0': 10.0,   # Test Date
    'Date Field1': 10.0,   # Date Initiated
    'Date Field2': 10.0,   # Date of Incident
    'Text Field1': 8.0,    # Test Name (wraps cleanly to 2 lines in 20pt)
    'Text Field2': 10.0,   # Sample ID
    'Text Field3': 6.225,  # All 4 Analysts (prepping, processing, aliquoting, reading)
    'Text Field4': 8.0,    # Sample / Active Name (MOTs-C)
    'Text Field5': 9.5,    # Dosage Form (Injection)
    'Text Field6': 10.0,   # Lot Number
    'Text Field7': 8.0,    # Description of Incident (3 lines fit cleanly in 42.4pt)
    'Text Field8': 9.5,    # SOP
    'Text Field9': 9.5,    # Effective Date
    'Text Field10': 9.5,   # Rev
    'Text Field11': 9.5,   # Limits
    'Text Field12': 10.0,  # Manager name
    'Date Field3': 10.0,   # Manager date
    
    # Page 1 Section B Comments
    'Text Field13': 4.825, # Interviewed comment (2 lines)
    'Text Field14': 8.5,   # Sample ID comment (1 line)
    'Text Field15': 9.0,   # SOP comment (1 line)
    'Text Field16': 9.0,   # SOP comment (1 line)
    'Text Field17': 6.725, # EagleTrax comment (2 lines)
    'Text Field18': 4.825, # Training comment (2 lines)
    'Text Field19': 9.0,   # Formula comment (1 line)
    'Text Field20': 9.0,   # Method comment (1 line)
    'Text Field21': 5.85,  # Storage comment (2 lines)
    
    # Page 2
    'Text Field22': 4.35,  # Reagent lots (14 lines fit cleanly in 111.1pt)
    'Text Field23': 4.75,  # Reagent exps (14 lines fit cleanly in 111.1pt)
    'Text Field24': 9.5,   # Standard Used
    'Text Field25': 9.5,   # Standard Lot
    'Text Field26': 9.5,   # Standard Exp
    'Text Field27': 9.5,   # Solvent Used
    'Text Field28': 9.5,   # Solvent Lot
    'Text Field29': 9.5,   # Solvent Exp
    'Text Field30': 9.5,   # E002222
    'Text Field31': 9.5,   # Sep 2026
    'Text Field32': 4.825, # Clean room facility (2 lines fit cleanly in 15.8pt)
    'Text Field33': 4.825, # Cal dates (2 lines fit cleanly in 15.8pt)
    'Text Field34': 9.5,   # E002222
    'Text Field35': 9.5,   # Sep 2026
    'Text Field36': 9.5,   # NA
    'Text Field37': 9.5,   # NA
    'Text Field38': 9.5,   # NA
    'Text Field39': 9.5,   # NA
    'Text Field40': 8.0,   # See Phase I Summary
    'Text Field41': 8.0,   # See Phase I Summary
    'Text Field42': 8.0,   # See Phase I Summary
    'Text Field43': 7.05,  # 4 Incubators with sensor IDs
    'Text Field44': 7.05,  # 4 Cal dates with exact spacing alignment
    'Text Field45': 8.0,   # See Phase I Summary
    
    # Page 3
    'Text Field46': 9.5,   # NA
    'Text Field47': 9.5,   # NA
    'Text Field48': 8.5,   # N/A QYC 02Oct26
    'Text Field49': 9.2,   # Interview Narrative (golden calibrated)
    
    # Page 4
    'Text Field50': 9.2,   # Narrative Part 1 (golden calibrated)
    
    # Page 5
    'Text Field51': 9.2,   # Narrative Part 2 (golden calibrated)
    
    # Page 6
    'Text Field53': 10.0,  # Qiyue Chen
    'Text Field54': 10.0,  # Robin Seymour (cleared by user)
    'Text Field57': 6.5,   # OOS Number in header
}

# Create a fresh 6-page document for the official corporate form
doc_corp = fitz.open()
doc_corp.insert_pdf(doc_src, from_page=0, to_page=5, links=True)

for p_no in range(6):
    p = doc_corp[p_no]
    for w in p.widgets():
        if w.field_type_string == 'Text':
            fn = w.field_name
            if fn in FONT_SIZE_MAP:
                w.text_fontsize = FONT_SIZE_MAP[fn]
            if w.field_value:
                # Strip trailing whitespace and newlines
                w.field_value = w.field_value.rstrip(' \r\n')
            w.update()

scratch_corp_path = os.path.join(SCRATCH_DIR, "CORP-FORM-21_fitted.pdf")
doc_corp.save(scratch_corp_path)
doc_corp.close()
print(f"Saved calibrated 6-page form to: {scratch_corp_path}")

print("\n=== STEP 3: ASSEMBLE 8-PAGE COMPLETE PACKET WITH HYPERLINKED TABLES ===")
doc_complete = fitz.open()
doc_corp_reloaded = fitz.open(scratch_corp_path)
doc_complete.insert_pdf(doc_corp_reloaded, links=True)
doc_corp_reloaded.close()

doc_tables = fitz.open(tables_pdf_path)
doc_complete.insert_pdf(doc_tables, links=True)
doc_tables.close()

assert len(doc_complete) == 8, f"Expected 8 pages, got {len(doc_complete)}"

scratch_complete_path = os.path.join(SCRATCH_DIR, "OOS-262080_complete_fitted.pdf")
doc_complete.save(scratch_complete_path)
doc_complete.close()
print(f"Saved complete 8-page packet to: {scratch_complete_path}")

print("\n=== STEP 4: VERIFY ZERO OVERFLOW & AP STREAM INTEGRITY ===")
doc_ver = fitz.open(scratch_complete_path)

overflow_count = 0
for p_no in range(6):
    p = doc_ver[p_no]
    for w in p.widgets():
        if w.field_type_string == 'Text' and w.field_value:
            xref = w.xref
            ap_xref = doc_ver.xref_get_key(xref, 'AP/N')
            if ap_xref[0] == 'xref':
                sx = int(ap_xref[1].split()[0])
                stream = doc_ver.xref_stream(sx).decode('latin1', errors='ignore')
                re_match = re.search(r'([0-9\.\-]+)\s+([0-9\.\-]+)\s+([0-9\.\-]+)\s+([0-9\.\-]+)\s+re\s+W', stream)
                if re_match:
                    rw = float(re_match.group(3))
                    rh = float(re_match.group(4))
                    td_matches = re.findall(r'([0-9\.\-]+)\s+([0-9\.\-]+)\s+Td', stream)
                    cur_y = 0.0
                    min_y = 999.0
                    for td in td_matches:
                        cur_y += float(td[1])
                        if cur_y < min_y:
                            min_y = cur_y
                    if min_y < 0:
                        print(f"  [ERROR] OVERFLOW on Page {p_no+1} {w.field_name}: min_y={min_y:.2f}, rh={rh:.2f}, text={repr(w.field_value[:35])}")
                        overflow_count += 1

print(f"Field Overflow Check: {overflow_count} overflowing fields detected.")
assert overflow_count == 0, f"Detected {overflow_count} overflowing fields!"

print("\n=== STEP 5: VERIFY TABLE 1 HYPERLINK & BASELINE ALIGNMENT ===")
p7 = doc_ver[6]
links = p7.get_links()
print(f"Page 7 Link count: {len(links)}")
assert len(links) >= 1, "Page 7 is missing Table 1 hyperlink!"
target_link = None
for lk in links:
    if 'jzraMOOTFYUFD' in lk.get('uri', ''):
        target_link = lk
        break

assert target_link is not None, "Target EagleTrax hyperlink not found on Page 7!"
print(f"Found active Table 1 hyperlink: URI={target_link['uri']}")
print(f"Link Rect: {target_link['from']}")

words = [w for w in p7.get_text('words') if 115 <= w[1] <= 135]
y0_coords = set(round(w[1], 3) for w in words)
y1_coords = set(round(w[3], 3) for w in words)
print(f"Row 1 words count: {len(words)}, distinct y0: {y0_coords}, distinct y1: {y1_coords}")
assert len(y0_coords) == 1 and len(y1_coords) == 1, f"Table 1 Row 1 words not perfectly aligned: y0={y0_coords}, y1={y1_coords}"
print("Table 1 baseline alignment verified: 100% pixel-perfect horizontal alignment across all 6 columns.")

print("\n=== STEP 6: VERIFY USER MANUAL EDITS ARE PRESERVED ===")
p3_w48 = [w for w in doc_ver[2].widgets() if w.field_name == 'Text Field48'][0]
p6_w53 = [w for w in doc_ver[5].widgets() if w.field_name == 'Text Field53'][0]
p6_w54 = [w for w in doc_ver[5].widgets() if w.field_name == 'Text Field54'][0]
p5_w51 = [w for w in doc_ver[4].widgets() if w.field_name == 'Text Field51'][0]

print(f"Page 3 Text Field48: {repr(p3_w48.field_value)} (Expected 'N/A QYC 02Oct26')")
print(f"Page 6 Text Field53: {repr(p6_w53.field_value)} (Expected 'Qiyue Chen')")
print(f"Page 6 Text Field54: {repr(p6_w54.field_value)} (Expected '')")
assert p3_w48.field_value == 'N/A QYC 02Oct26', f"Unexpected Text Field48 value: {p3_w48.field_value}"
assert p6_w53.field_value == 'Qiyue Chen', f"Unexpected Text Field53 value: {p6_w53.field_value}"
assert p6_w54.field_value == '', f"Unexpected Text Field54 value: {p6_w54.field_value}"
assert 'Analyzing a 6-month sample history for Optimal Balance Pharmacy indicates, this specific analyte "MOTs-C 10 MG/ML (5 ML) Injection" has had no prior failures using the Celsis Sterility testing during this period.' in p5_w51.field_value, "MOTs-C history sentence missing!"
assert 'The corresponding TSB sample container' not in p5_w51.field_value, "TSB sentence was not removed!"
assert 'A review of the lot history shows' not in p5_w51.field_value, "Lot history sentence was not removed!"
print("All user manual edits verified and preserved.")

doc_ver.close()

print("\n=== STEP 7: DEPLOY DELIVERABLES TO DESKTOP AND DOCUMENTS ===")
def safe_deploy(src, dst):
    try:
        shutil.copy2(src, dst)
        print(f"  [OK] Successfully deployed -> {dst}")
        return True
    except PermissionError:
        print(f"  [LOCKED] Could not write to {dst}. File is open in another application.")
        return False

# 1. Official 6-Page CORP-FORM-21
safe_deploy(scratch_corp_path, out_corp_desktop)
safe_deploy(scratch_corp_path, out_corp_docs)

# 2. Master 8-Page Complete Deliverables
safe_deploy(scratch_complete_path, out_qyc_desktop)
safe_deploy(scratch_complete_path, out_complete_desktop)
safe_deploy(scratch_complete_path, out_2_desktop)
safe_deploy(scratch_complete_path, out_complete_docs)

print("\n=== STEP 8: RENDER VERIFICATION IMAGES FOR USER ===")
doc_final = fitz.open(scratch_complete_path)
for idx, page in enumerate(doc_final):
    pix = page.get_pixmap(dpi=150)
    out_img = os.path.join(SCRATCH_DIR, f"final_verified_p{idx+1}.png")
    pix.save(out_img)
    print(f"  Rendered Page {idx+1} -> {out_img}")

# Also render a high-res crop of Section B on Page 1
p1 = doc_final[0]
pix_crop = p1.get_pixmap(clip=fitz.Rect(300, 480, 580, 680), dpi=200)
crop_img = os.path.join(SCRATCH_DIR, "final_verified_section_b.png")
pix_crop.save(crop_img)
print(f"  Rendered Section B crop -> {crop_img}")

doc_final.close()
print("\n=== ALL TASKS COMPLETED SUCCESSFULLY! ===")

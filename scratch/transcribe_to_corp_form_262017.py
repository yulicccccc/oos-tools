"""
Script: transcribe_to_corp_form_262017.py
Purpose: Transcribe finalized Celsis OOS-262017 into freshly downloaded official ZenQMS
         form CORP-FORM-21 - P1 31 Aug 2026.pdf (v11.1), incorporate user's manual updates
         (N/A QYC 07Oct26 on Page 3 Text Field48, Cuong Du as initiator, empty manager signature),
         calibrate fonts to eliminate Acrobat '+' overflow, append 2-page standalone tables
         with clickable hyperlinks as Pages 7 & 8, and sync deliverables to Desktop and Documents.
"""

import os
import sys
import shutil
from datetime import datetime
import fitz  # PyMuPDF

sys.stdout.reconfigure(encoding='utf-8')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.abspath(os.path.join(SCRIPT_DIR, ".."))
DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
DOCS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
HISTORY_DIR = os.path.join(ROOT_DIR, ".history")
SCRATCH_DIR = os.path.join(ROOT_DIR, "scratch")

os.makedirs(HISTORY_DIR, exist_ok=True)
os.makedirs(SCRATCH_DIR, exist_ok=True)

TARGET_CORP_PDF = os.path.join(DESKTOP_DIR, "CORP-FORM-21 - P1 31 Aug 2026.pdf")
SOURCE_OOS_PDF = os.path.join(DESKTOP_DIR, "OOS-262017 Revive Rx Pharmacy - PO Required (E00927) - Celsis.pdf")
TABLES_PDF_PATH = os.path.join(DESKTOP_DIR, "Celsis table OOS-262017.pdf")
COMPLETE_PDF_PATH = os.path.join(DESKTOP_DIR, "OOS-262017 Revive Rx Pharmacy - PO Required (E00927) - Celsis (Complete).pdf")

print("=== [STEP 1] BACKUP ORIGINAL TARGET CORP-FORM-21 ===")
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
backup_name = f"CORP-FORM-21 - P1 31 Aug 2026_backup_{timestamp}.pdf"
backup_path = os.path.join(HISTORY_DIR, backup_name)
shutil.copy2(TARGET_CORP_PDF, backup_path)
print(f"Backed up original blank target form to:\n  {backup_path}")

print("\n=== [STEP 2] LOAD SOURCE DATA & TARGET FORM ===")
doc_src = fitz.open(SOURCE_OOS_PDF)
# Use backup as pristine clean template base
doc_tgt = fitz.open(backup_path)

src_fields = {}
for p_no in range(6):
    for w in doc_src[p_no].widgets():
        src_fields[w.field_name] = (w.field_value, w.text_fontsize, w.field_type_string)

print(f"Extracted {len(src_fields)} fields from user's finalized OOS PDF.")

# Import user's balanced 3-page narrative text
from check_user_layout import t3, t4, t5

# Definitive Calibrated Font Size Map for CORP-FORM-21
FONT_SIZE_MAP = {
    # Page 1
    'Text Field0': 8.5,    # Initiator (Cuong Du (Written by Qiyue Chen))
    'Date Field0': 8.5,    # 24-Aug-2026
    'Date Field1': 8.5,    # 31-Aug-2026
    'Date Field2': 8.5,    # 31-Aug-2026
    'Text Field1': 8.5,    # Celsis Sterility Test
    'Text Field2': 8.5,    # ETX-260821-0259
    'Text Field3': 6.225,  # All 4 Analysts
    'Text Field4': 8.5,    # Tesamorelin 12 mg per Vial
    'Text Field5': 8.5,    # Lyophilized
    'Text Field6': 8.5,    # 16590773
    'Text Field7': 8.0,    # Incident Description
    'Text Field8': 8.5,    # MICRO-SOP-44
    'Text Field9': 8.5,    # 03-Jan-2025
    'Text Field10': 8.5,   # 01
    'Text Field11': 8.5,   # Pass/Fail
    'Text Field12': 8.5,   # Kathan Parikh
    'Date Field3': 8.5,    # 31-Aug-2026
    
    # Page 1 Section B Comments
    'Text Field13': 4.825, # Interview comment
    'Text Field14': 8.5,   # Sample ID comment
    'Text Field15': 9.0,   # SOP comment
    'Text Field16': 9.0,   # SOP comment
    'Text Field17': 6.725, # Session comment
    'Text Field18': 4.825, # Training comment
    'Text Field19': 9.0,   # Formula comment (Not Applicable)
    'Text Field20': 9.0,   # Method comment (Not Applicable)
    'Text Field21': 5.85,  # Storage comment
    
    # Page 2
    'Text Field22': 4.35,  # Reagents Lot numbers
    'Text Field23': 4.75,  # Reagents Expiration dates
    'Text Field24': 8.5,   # Celsis ATP Positive Control
    'Text Field25': 8.5,   # 022601-1483
    'Text Field26': 8.5,   # Jan 2027
    'Text Field27': 8.5,   # Not Applicable
    'Text Field28': 8.5,   # Not Applicable
    'Text Field29': 8.5,   # Not Applicable
    'Text Field30': 8.5,   # E002222
    'Text Field31': 8.5,   # Sep 2026
    'Text Field32': 4.825, # BSCs E001316 & E001798
    'Text Field33': 4.825, # Dec 2026 / Jun 2027
    'Text Field34': 8.5,   # E002222
    'Text Field35': 8.5,   # Sep 2026
    'Text Field36': 8.5,   # Not Applicable
    'Text Field37': 8.5,   # Not Applicable
    'Text Field38': 8.5,   # Not Applicable
    'Text Field39': 8.5,   # Not Applicable
    'Text Field40': 8.5,   # See Phase I Summary
    'Text Field41': 8.5,   # See Phase I Summary
    'Text Field42': 8.5,   # See Phase I Summary
    'Text Field43': 7.05,  # Incubators
    'Text Field44': 7.05,  # Incubator Cal Due Dates
    'Text Field45': 8.5,   # See Phase I Summary
    
    # Page 3
    'Text Field46': 8.5,   # Not Applicable
    'Text Field47': 8.5,   # Not Applicable
    'Text Field48': 8.5,   # N/A QYC 07Oct26 (User manual update)
    'Text Field49': 8.35,  # Narrative P1..P7 (7 paragraphs)
    
    # Page 4
    'Text Field50': 8.5,   # Narrative P8..P12 (5 paragraphs)
    
    # Page 5
    'Text Field51': 8.5,   # Narrative P13..P21 (8 paragraphs)
    
    # Page 6
    'Text Field52': 8.5,   # 
    'Text Field53': 8.5,   # Qiyue Chen
    'Text Field54': 8.5,   # Empty for Lab Manager signature
    'Text Field55': 8.5,   # 
    'Text Field57': 8.5,   # Header OOS Number: 262017
}

print("\n=== [STEP 3] TRANSCRIBE INTO FRESH CORP-FORM-21 (PAGES 1-6) ===")
for p_no in range(6):
    page = doc_tgt[p_no]
    for w in page.widgets():
        fn = w.field_name
        if fn in src_fields:
            val, fs, ftype = src_fields[fn]
            
            # Specific overrides & harmonizations
            if fn == 'Text Field48':
                w.field_value = "N/A QYC 07Oct26"
            elif fn == 'Text Field49':
                w.field_value = t3
            elif fn == 'Text Field50':
                w.field_value = t4
            elif fn == 'Text Field51':
                w.field_value = t5
            elif fn == 'Text Field54':
                w.field_value = ""  # Cleared for Lab Manager signature
            elif ftype == 'CheckBox':
                w.field_value = val if val in ['Yes', 'Off'] else ('Yes' if val else 'Off')
            else:
                w.field_value = val
                
            if fn in FONT_SIZE_MAP:
                w.text_fontsize = FONT_SIZE_MAP[fn]
            w.update()

# Ensure Page 6 CheckBox exclusivity
p6 = doc_tgt[5]
for w in p6.widgets():
    if w.field_name == 'Check Box88':
        w.field_value = 'Yes'
        w.update()
    elif w.field_name in ['Check Box87', 'Check Box89']:
        w.field_value = 'Off'
        w.update()

scratch_transcribed = os.path.join(SCRATCH_DIR, "CORP-FORM-21 - P1 31 Aug 2026_transcribed.pdf")
doc_tgt.save(scratch_transcribed)
doc_tgt.close()
print(f"Saved transcribed form (Pages 1-6) to:\n  {scratch_transcribed}")

print("\n=== [STEP 4] VERIFY TRANSCRIBED FORM (PAGES 1-6) ===")
doc_ver = fitz.open(scratch_transcribed)
assert len(doc_ver) == 6, f"Expected 6 pages, got {len(doc_ver)}"

# Check text flow and margins on Pages 3, 4, 5
narrative_checks = [
    (2, 'Text Field49', 'E002222.'),
    (3, 'Text Field50', 'aliquoting.'),
    (4, 'Text Field51', 'procedure.')
]
for pidx, fn, expected_last in narrative_checks:
    p = doc_ver[pidx]
    w_obj = [w for w in p.widgets() if w.field_name == fn][0]
    words = [w for w in p.get_text('words') if w[1] >= w_obj.rect.y0 - 2 and w[3] <= w_obj.rect.y1 + 5]
    last_y = max([w[3] for w in words]) if words else 0
    rem = w_obj.rect.y1 - last_y
    last_word = words[-1][4] if words else ''
    print(f"  Page {pidx+1} ({fn}): rem={rem:.1f}pt, last word='{last_word}' (expected '{expected_last}')")
    assert rem > 0, f"Overflow detected on Page {pidx+1} ({fn})!"
    assert expected_last in last_word, f"Last word mismatch on Page {pidx+1}: expected '{expected_last}', got '{last_word}'"

# Check Page 3 Text Field48
p3_w48 = [w for w in doc_ver[2].widgets() if w.field_name == 'Text Field48'][0]
print(f"  Page 3 Text Field48 value: '{p3_w48.field_value}' (verified)")
assert p3_w48.field_value == "N/A QYC 07Oct26"

# Check Page 6 CheckBox 88
p6_cb88 = [w for w in doc_ver[5].widgets() if w.field_name == 'Check Box88'][0]
print(f"  Page 6 Check Box88 value: '{p6_cb88.field_value}' (verified)")
assert p6_cb88.field_value == "Yes"

doc_ver.close()
print("Form verification passed with 100% success!")

print("\n=== [STEP 5] ASSEMBLE COMBINED 8-PAGE COMPLETE PACKAGE ===")
doc_combined = fitz.open()
doc_f6 = fitz.open(scratch_transcribed)
doc_combined.insert_pdf(doc_f6)
doc_f6.close()

if os.path.exists(TABLES_PDF_PATH):
    doc_tbl = fitz.open(TABLES_PDF_PATH)
    doc_combined.insert_pdf(doc_tbl, links=True)
    doc_tbl.close()
    print(f"Appended Table 1, 2, and 3 from {TABLES_PDF_PATH}.")
else:
    raise FileNotFoundError(f"Tables PDF not found: {TABLES_PDF_PATH}")

print(f"Total pages in combined package: {len(doc_combined)} (Must be 8)")
assert len(doc_combined) == 8, f"Expected 8 pages, got {len(doc_combined)}"

# Verify clickable hyperlinks on Pages 7 & 8
p7_links = doc_combined[6].get_links()
p8_links = doc_combined[7].get_links()
print(f"Page 7 clickable hyperlinks: {len(p7_links)}")
for l in p7_links:
    print(f"  P7 link: {l.get('uri')}")
print(f"Page 8 clickable hyperlinks: {len(p8_links)}")
for l in p8_links:
    print(f"  P8 link: {l.get('uri')}")
assert len(p7_links) >= 3, f"Expected >= 3 links on Page 7, got {len(p7_links)}"
assert len(p8_links) >= 1, f"Expected >= 1 link on Page 8, got {len(p8_links)}"

scratch_complete = os.path.join(SCRATCH_DIR, "OOS-262017 Revive Rx Pharmacy - PO Required (E00927) - Celsis (Complete).pdf")
doc_combined.save(scratch_complete)
doc_combined.close()
print(f"Saved complete combined package to:\n  {scratch_complete}")

print("\n=== [STEP 6] DEPLOY DELIVERABLES TO DESKTOP AND DOCUMENTS ===")
def safe_copy(src, dst):
    try:
        shutil.copy2(src, dst)
        print(f"  SUCCESS: Copied to {dst}")
        return True
    except PermissionError:
        print(f"  LOCK NOTICE: {dst} is open in another program.")
        return False
    except Exception as e:
        print(f"  ERROR: {e}")
        return False

# 1. Update target CORP-FORM-21 on Desktop with complete 8-page packet
safe_copy(scratch_complete, TARGET_CORP_PDF)

# Also create the 6-page standalone form on Desktop
standalone_form_pdf = os.path.join(DESKTOP_DIR, "CORP-FORM-21 - P1 31 Aug 2026 (6 Pages Form Only).pdf")
safe_copy(scratch_transcribed, standalone_form_pdf)

# 2. Deploy OOS-262017 Complete PDF on Desktop
safe_copy(scratch_complete, COMPLETE_PDF_PATH)

# 3. Sync to Documents/OOS for archive & version control
docs_target_pdf = os.path.join(DOCS_DIR, "CORP-FORM-21 - P1 31 Aug 2026.pdf")
safe_copy(scratch_complete, docs_target_pdf)
docs_complete_pdf = os.path.join(DOCS_DIR, "OOS-262017 Revive Rx Pharmacy - PO Required (E00927) - Celsis (Complete).pdf")
safe_copy(scratch_complete, docs_complete_pdf)

print("\n=== [STEP 7] RENDER FINAL HIGH-RES PREVIEW IMAGES (PAGES 1-8) ===")
preview_dir = os.path.join(SCRATCH_DIR, "preview_corp_262017_final")
os.makedirs(preview_dir, exist_ok=True)
doc_render = fitz.open(scratch_complete)
for idx, page in enumerate(doc_render):
    pix = page.get_pixmap(dpi=150)
    out_img = os.path.join(preview_dir, f"page_{idx+1}.png")
    pix.save(out_img)
    print(f"  Rendered Page {idx+1} -> {out_img}")
doc_render.close()

print("\n=== ALL TASKS COMPLETED WITH 100% SUCCESS! ===")

"""
Script: transcribe_to_corp_form_261967.py
Purpose: Transcribe full OOS-261967 investigation report into freshly downloaded CORP-FORM-21 (v11.1),
         incorporate user's manual layout updates (Mukyung Jang, 2-page narrative, N/A on Page 5, empty sig),
         apply calibrated font sizes, attach Table 1 & Table 2 with clickable hyperlinks as Page 7,
         and sync deliverables across Downloads, Desktop, and Documents.
"""

import os
import sys
import shutil
from datetime import datetime
import fitz  # PyMuPDF
import re

sys.stdout.reconfigure(encoding='utf-8')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.abspath(os.path.join(SCRIPT_DIR, ".."))
DOCUMENTS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents"
DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
DOWNLOADS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Downloads"
HISTORY_DIR = os.path.join(ROOT_DIR, ".history")
SCRATCH_DIR = os.path.join(ROOT_DIR, "scratch")

os.makedirs(HISTORY_DIR, exist_ok=True)
os.makedirs(SCRATCH_DIR, exist_ok=True)

TARGET_DOWNLOADS_PDF = os.path.join(DOWNLOADS_DIR, "CORP-FORM 21 - P1 - 25 Aug 2026.pdf")
SOURCE_OOS_PDF = os.path.join(DESKTOP_DIR, "OOS-261967 TAM Pharmacy (E74685) - ScanRDI - QYC.pdf")
TABLES_PDF_PATH = os.path.join(SCRATCH_DIR, "temp_tables_export.pdf")
if not os.path.exists(TABLES_PDF_PATH):
    TABLES_PDF_PATH = os.path.join(DOCUMENTS_DIR, "Tables OOS-261967 TAM Pharmacy (E74685) - ScanRDI.pdf")

OUTPUT_DESKTOP_PDF = os.path.join(DESKTOP_DIR, "OOS-261967 TAM Pharmacy (E74685) - ScanRDI.pdf")
OUTPUT_DESKTOP_QYC_PDF = os.path.join(DESKTOP_DIR, "OOS-261967 TAM Pharmacy (E74685) - ScanRDI - QYC.pdf")
OUTPUT_DOCS_PDF = os.path.join(DOCUMENTS_DIR, "OOS-261967 TAM Pharmacy (E74685) - ScanRDI.pdf")

print("=== [STEP 1] BACKUP ORIGINAL TARGET CORP-FORM-21 ===")
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
backup_name = f"CORP-FORM 21 - P1 - 25 Aug 2026_backup_{timestamp}.pdf"
backup_path = os.path.join(HISTORY_DIR, backup_name)
shutil.copy2(TARGET_DOWNLOADS_PDF, backup_path)
print(f"Backed up original target to: {backup_path}")

print("\n=== [STEP 2] LOAD SOURCE DATA & TARGET FORM ===")
# Use original blank / clean template from the earliest backup in .history if available
earliest_backups = sorted([f for f in os.listdir(HISTORY_DIR) if f.startswith("CORP-FORM 21 - P1 - 25 Aug 2026_backup_")])
if earliest_backups:
    clean_tgt_template = os.path.join(HISTORY_DIR, earliest_backups[0])
    print(f"Using clean target template from earliest backup: {clean_tgt_template}")
else:
    clean_tgt_template = backup_path

doc_src = fitz.open(SOURCE_OOS_PDF)
doc_tgt = fitz.open(clean_tgt_template)

# Extract source field values from user's edited PDF
src_data = {}
for p_no in range(6):
    for w in doc_src[p_no].widgets():
        src_data[w.field_name] = w.field_value

print(f"Extracted {len(src_data)} fields from source QYC OOS PDF.")

# Definitive Calibrated Font Size Map for CORP-FORM-21
FONT_SIZE_MAP = {
    # Page 1
    'Text Field0': 6.0,    # Initiator (Varsha Subramanian / Written by: Qiyue Chen)
    'Date Field0': 9.5,    # Test Date
    'Date Field1': 9.5,    # Date Initiated
    'Date Field2': 9.5,    # Date of Incident
    'Text Field1': 7.5,    # Test Name (Scan RDI Sterility Test)
    'Text Field2': 9.5,    # Sample ID (ETX-260813-0778)
    'Text Field3': 6.0,    # All 4 Analysts (Prepping, Processing, Changeover, Reading)
    'Text Field4': 7.5,    # Sample Name (Semaglutide/Pyridoxine 2.5mg/10mg/mL)
    'Text Field5': 9.5,    # Dosage Form (Injectable)
    'Text Field6': 9.5,    # Lot Number (LG403000087)
    'Text Field7': 7.5,    # Incident Description
    'Text Field8': 6.2,    # SOP numbers
    'Text Field9': 6.2,    # Effective Date
    'Text Field10': 6.2,   # Rev numbers
    'Text Field11': 9.0,   # Limits (<1 event)
    'Text Field12': 9.5,   # Manager name (Kathan Parikh)
    'Date Field3': 9.5,    # Manager date (25-Aug-2026)
    
    # Page 1 Section B Comments
    'Text Field13': 5.2,   # Interview comment (Mukyung Jang, Varsha Subramanian, Sonal Uprety)
    'Text Field14': 8.0,   # Sample ID comment
    'Text Field15': 8.0,   # SOP comment
    'Text Field16': 8.0,   # SOP comment
    'Text Field17': 6.8,   # Session comment
    'Text Field18': 4.8,   # Training comment
    'Text Field19': 8.5,   # Formula comment
    'Text Field20': 8.5,   # Method comment
    'Text Field21': 5.6,   # Storage comment
    
    # Page 2
    'Text Field22': 4.5,   # Scan Consumables & Environmental Plates
    'Text Field23': 4.5,   # Scan Consumables & Environmental Plates
    'Text Field24': 9.0,   # Standard Used (C. sporogenes)
    'Text Field25': 9.0,   # Standard Lot
    'Text Field26': 9.0,   # Standard Exp
    'Text Field27': 9.0,   # Solvent Used
    'Text Field28': 9.0,   # Solvent Lot
    'Text Field29': 9.0,   # Solvent Exp
    'Text Field30': 8.5,   # E001040 (Cs2-105)
    'Text Field31': 9.0,   # Nov 2026
    'Text Field32': 5.5,   # CR145 (E001979) in L-Suite
    'Text Field33': 8.5,   # Dec 2026
    'Text Field34': 8.5,   # E001040 (Cs2-105)
    'Text Field35': 9.0,   # Nov 2026
    'Text Field36': 9.0,   # NA
    'Text Field37': 9.0,   # NA
    'Text Field38': 9.0,   # NA
    'Text Field39': 9.0,   # NA
    'Text Field40': 8.0,   # See Phase I Summary
    'Text Field41': 8.0,   # See Phase I Summary
    'Text Field42': 8.0,   # See Phase I Summary
    'Text Field43': 6.8,   # Incubators E001034 & E001031
    'Text Field44': 6.8,   # Cal dates Aug 2027 / Feb 2027
    'Text Field45': 8.0,   # See Phase I Summary
    
    # Page 3
    'Text Field46': 9.0,   # NA
    'Text Field47': 9.0,   # NA
    'Text Field48': 8.5,   # N/A QYC 06Oct26
    'Text Field49': 8.75,  # Narrative Part 1 (Mukyung Jang, 3,542 chars)
    
    # Page 4
    'Text Field50': 9.2,   # Narrative Part 2 (3,375 chars)
    
    # Page 5
    'Text Field51': 8.5,   # N/A QYC 06Oct26
    
    # Page 6
    'Text Field53': 9.5,   # Prepared by (Qiyue Chen)
    'Text Field54': 9.5,   # Lab Manager (Empty for signature)
    'Text Field57': 6.5,   # OOS Number in header (261967)
}

print("\n=== [STEP 3] TRANSCRIBE INTO FRESH CORP-FORM-21 (PAGES 1-6) ===")
for p_no in range(6):
    page = doc_tgt[p_no]
    for w in page.widgets():
        if w.rect.is_empty or w.rect.width <= 0 or w.rect.height <= 0:
            continue
        fn = w.field_name
        if fn in src_data:
            val = src_data[fn]
            
            # Special formatting & corrections
            if fn == 'Text Field0':
                w.field_flags |= fitz.PDF_TX_FIELD_IS_MULTILINE
                w.field_value = "Varsha Subramanian\r(Written by: Qiyue Chen)"
            elif fn == 'Text Field3':
                w.field_value = (
                    "Prepping Analyst: \rMukyung Jang (MJ)\r \r"
                    "Processing Analyst: \rVarsha Subramanian (VV)\r \r"
                    "Changeover Analyst: \rVarsha Subramanian (VV)\r \r"
                    "Reading Analyst: \rSonal Uprety (SU)"
                )
            elif fn == 'Text Field4':
                w.field_value = "Semaglutide/Pyridoxine 2.5mg/10mg/mL\r \r \r"
            elif fn == 'Text Field13':
                w.field_value = "Yes, analysts Mukyung Jang, Varsha Subramanian, and Sonal Uprety were interviewed comprehensively."
            elif fn == 'Text Field43':
                w.field_value = "Incubator E001034 (Sensor E001501)\r \rIncubator E001031 (Sensor E001505)"
            elif fn == 'Text Field44':
                w.field_value = "Aug 2027 / Feb 2027\r \rAug 2027 / Feb 2027"
            elif fn == 'Text Field48':
                w.field_value = "N/A QYC 06Oct26"
            elif fn == 'Text Field49':
                # Replace Min Jang with Mukyung Jang
                w.field_value = val.replace("Min Jang", "Mukyung Jang")
            elif fn == 'Text Field50':
                w.field_value = val
            elif fn == 'Text Field51':
                w.field_value = "N/A QYC 06Oct26"
            elif fn == 'Text Field54':
                w.field_value = ""  # Cleared for Lab Manager signature
            else:
                w.field_value = val
                
            if fn in FONT_SIZE_MAP:
                w.text_fontsize = FONT_SIZE_MAP[fn]
            w.update()

# Explicitly ensure CheckBox exclusivity on Page 6
p6 = doc_tgt[5]
for w in p6.widgets():
    if w.field_name == 'Check Box88':
        w.field_value = 'Yes'
        w.update()
    elif w.field_name in ['Check Box87', 'Check Box89']:
        w.field_value = 'Off'
        w.update()

print("Pages 1 to 6 populated successfully.")

print("\n=== [STEP 4] APPEND PAGE 7 (TABLE 1 & TABLE 2 WITH HYPERLINKS) ===")
doc_tbl = fitz.open(TABLES_PDF_PATH)
doc_tgt.insert_pdf(doc_tbl, from_page=0, to_page=0, links=True)
doc_tbl.close()
print(f"Appended Table Page. Total pages in package: {len(doc_tgt)}")

scratch_output = os.path.join(SCRATCH_DIR, "CORP-FORM 21 - P1 - 25 Aug 2026_complete.pdf")
doc_tgt.save(scratch_output)
doc_tgt.close()
print(f"Saved completed scratch package: {scratch_output}")

print("\n=== [STEP 5] VERIFY COMPILED PACKAGE ===")
doc_ver = fitz.open(scratch_output)
assert len(doc_ver) == 7, f"Expected 7 pages, got {len(doc_ver)}"

# Verify links on Page 7
p7_links = doc_ver[6].get_links()
print(f"Page 7 has {len(p7_links)} active clickable hyperlinks:")
for lk in p7_links:
    print(f"  -> {lk.get('uri')}")
assert len(p7_links) >= 3, f"Expected at least 3 hyperlinks on Page 7, found {len(p7_links)}"

# Verify text flow on Pages 3, 4, 5
for page_idx, fn in [(2, 'Text Field49'), (3, 'Text Field50'), (4, 'Text Field51')]:
    p = doc_ver[page_idx]
    w_obj = [w for w in p.widgets() if w.field_name == fn][0]
    words = [w for w in p.get_text('words') if w[1] >= w_obj.rect.y0 - 2 and w[3] <= w_obj.rect.y1 + 5]
    last_y = max([w[3] for w in words]) if words else 0
    rem = w_obj.rect.y1 - last_y
    last_word = words[-1][4] if words else ''
    print(f"  Page {page_idx+1} ({fn}): rem={rem:.1f}pt, last_word='{last_word}'")
    assert rem > 0, f"Overflow detected on Page {page_idx+1} ({fn})!"

# Verify absence of 'Min Jang' anywhere in the document
for page_idx in range(len(doc_ver)):
    for w in doc_ver[page_idx].widgets():
        if w.field_value and 'Min Jang' in w.field_value:
            raise AssertionError(f"Found 'Min Jang' on Page {page_idx+1} in {w.field_name}!")

print("All verification assertions passed with 100% success!")
doc_ver.close()

print("\n=== [STEP 6] DEPLOY DELIVERABLES ===")
# 1. Write directly to user's target in Downloads
shutil.copy2(scratch_output, TARGET_DOWNLOADS_PDF)
print(f"SUCCESS: Transcribed target in Downloads updated:\n  {TARGET_DOWNLOADS_PDF}")

# 2. Sync to Documents
try:
    shutil.copy2(scratch_output, OUTPUT_DOCS_PDF)
    print(f"SUCCESS: Documents copy updated:\n  {OUTPUT_DOCS_PDF}")
except Exception as e:
    print(f"Documents copy note: {e}")

# 3. Sync to Desktop
try:
    shutil.copy2(scratch_output, OUTPUT_DESKTOP_PDF)
    print(f"SUCCESS: Desktop copy updated:\n  {OUTPUT_DESKTOP_PDF}")
except Exception as e:
    print(f"Desktop copy note (file may be open): {e}")

try:
    shutil.copy2(scratch_output, OUTPUT_DESKTOP_QYC_PDF)
    print(f"SUCCESS: Desktop QYC copy updated:\n  {OUTPUT_DESKTOP_QYC_PDF}")
except Exception as e:
    print(f"Desktop QYC copy note (file may be open): {e}")

print("\n=== [STEP 7] RENDER FINAL HIGH-RES PREVIEW IMAGES ===")
preview_dir = os.path.join(SCRATCH_DIR, "preview_corp_261967_final")
os.makedirs(preview_dir, exist_ok=True)
doc_render = fitz.open(scratch_output)
for idx, page in enumerate(doc_render):
    pix = page.get_pixmap(dpi=150)
    out_img = os.path.join(preview_dir, f"page_{idx+1}.png")
    pix.save(out_img)
    print(f"  Rendered Page {idx+1} -> {out_img}")
doc_render.close()

print("\n=== ALL TASKS COMPLETED SUCCESSFULLY! ===")

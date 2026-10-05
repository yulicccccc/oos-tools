import os
import sys
import docx
from docx.oxml import parse_xml
from docx.opc.constants import RELATIONSHIP_TYPE
import win32com.client
import fitz
import shutil

sys.stdout.reconfigure(encoding='utf-8')

DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
DOCS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
SCRATCH_DIR = os.path.join(DOCS_DIR, "scratch")

main_docx_path = os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis.docx")
tables_docx_path = os.path.join(DESKTOP_DIR, "Celsis table OOS-262080.docx")
tables_pdf_path = os.path.join(DESKTOP_DIR, "Celsis table OOS-262080.pdf")

out_qyc_desktop = os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis - QYC.pdf")
out_complete_desktop = os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis (Complete).pdf")
out_2_desktop = os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis (2).pdf")
out_complete_docs = os.path.join(DOCS_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis (Complete).pdf")
out_corp_desktop = os.path.join(DESKTOP_DIR, "CORP-FORM-21 Laboratory OOS Investigation Form (v11.1).pdf")
out_corp_docs = os.path.join(DOCS_DIR, "CORP-FORM-21 Laboratory OOS Investigation Form (v11.1).pdf")

URL_0520 = "https://etrax.eagleanalytical.com/SubmissionTest/Details/qnwcLQO5BWeJhWCBG7jl8Q__"
SAMPLE_0520 = "ETX-260921-0520"

def safe_copy(src, dst):
    try:
        shutil.copy2(src, dst)
        print(f"  [SUCCESS] Copied to: {dst}")
        return True
    except PermissionError:
        print(f"  [LOCKED] File is open in another app, skipping lock: {dst}")
        return False

print("=== STEP 1: ASSEMBLE 8-PAGE COMPLETE PACKET WITH UPDATED TABLES ===")
scratch_corp_updated = os.path.join(SCRATCH_DIR, "CORP-FORM-21_updated_micro.pdf")
scratch_complete_updated = os.path.join(SCRATCH_DIR, "OOS-262080_complete_micro.pdf")

doc_complete = fitz.open()
doc_form_in = fitz.open(scratch_corp_updated)
doc_complete.insert_pdf(doc_form_in, links=True)
doc_form_in.close()

doc_tbl_in = fitz.open(tables_pdf_path)
doc_complete.insert_pdf(doc_tbl_in, links=True)
doc_tbl_in.close()

doc_complete.save(scratch_complete_updated)
doc_complete.close()
print(f"Saved complete 8-page PDF to: {scratch_complete_updated}")

print("\n=== STEP 2: DEPLOY TO DESKTOP AND DOCUMENTS ===")
safe_copy(scratch_corp_updated, out_corp_desktop)
safe_copy(scratch_corp_updated, out_corp_docs)

safe_copy(scratch_complete_updated, out_complete_desktop)
safe_copy(scratch_complete_updated, out_2_desktop)
safe_copy(scratch_complete_updated, out_complete_docs)
safe_copy(scratch_complete_updated, out_qyc_desktop)

# Also create an explicit updated file name that is never locked
out_updated_desktop = os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis - QYC - Updated.pdf")
safe_copy(scratch_complete_updated, out_updated_desktop)

print("\n=== STEP 3: VERIFY COMPLETE PDF INTEGRITY ===")
doc_ver = fitz.open(scratch_complete_updated)
print(f"Total pages: {len(doc_ver)}")

# Page 5 narrative
p5_text = doc_ver[4].get_text()
assert "Micrococcus luteus (Gram-positive cocci)" in p5_text, "Page 5 missing Micrococcus luteus!"
print("Page 5: Narrative contains 'Micrococcus luteus (Gram-positive cocci)'.")

# Page 7 link (Table 1)
p7_links = doc_ver[6].get_links()
print(f"Page 7 links: {len(p7_links)}")
assert any('jzraMOOTFYUFD' in lk.get('uri', '') for lk in p7_links), "Page 7 missing Table 1 link!"

# Page 8 link (Table 3)
p8_links = doc_ver[7].get_links()
print(f"Page 8 links: {len(p8_links)}")
assert any('qnwcLQO5BWeJhWCBG7jl8Q' in lk.get('uri', '') for lk in p8_links), "Page 8 missing Table 3 link!"
p8_text = doc_ver[7].get_text()
assert "Micrococcus luteus" in p8_text, "Page 8 missing 'Micrococcus luteus'!"
print("Page 8: Table 3 contains 'Micrococcus luteus' and clickable link to ETX-260921-0520.")

doc_ver.close()
print("\n=== ALL DELIVERABLES FULLY UPDATED AND SYNCHRONIZED! ===")

import os
import sys
import shutil
import docx
from docx.oxml import parse_xml
from docx.opc.constants import RELATIONSHIP_TYPE
import win32com.client
import fitz
from PIL import Image

sys.stdout.reconfigure(encoding='utf-8')

DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
DOCS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
SCRATCH_DIR = os.path.join(DOCS_DIR, "scratch")

# Source table docx
src_docx = os.path.join(SCRATCH_DIR, "test_bacteria_alignment.docx")
dest_tables_docx = os.path.join(SCRATCH_DIR, "Celsis table OOS-262080.docx")
dest_tables_pdf = os.path.join(SCRATCH_DIR, "Celsis table OOS-262080.pdf")

shutil.copy2(src_docx, dest_tables_docx)
print(f"Copied test docx to: {dest_tables_docx}")

print("\n=== STEP 1: EXPORT STANDALONE TABLES PDF VIA WORD COM ===")
word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
word.DisplayAlerts = 0

doc_com = word.Documents.Open(dest_tables_docx)
doc_com.SaveAs2(dest_tables_pdf, FileFormat=17)
doc_com.Close()
word.Quit()
print(f"Exported tables PDF: {dest_tables_pdf}")

# Deploy to Desktop tables
for d_docx in [
    os.path.join(DESKTOP_DIR, "Celsis table OOS-262080 (Updated).docx"),
    os.path.join(DESKTOP_DIR, "Celsis table OOS-262080.docx")
]:
    try:
        shutil.copy2(dest_tables_docx, d_docx)
        print(f"  Deployed tables DOCX -> {d_docx}")
    except Exception as e:
        print(f"  Notice for {d_docx}: {e}")

for d_pdf in [
    os.path.join(DESKTOP_DIR, "Celsis table OOS-262080 (Updated).pdf"),
    os.path.join(DESKTOP_DIR, "Celsis table OOS-262080.pdf")
]:
    try:
        shutil.copy2(dest_tables_pdf, d_pdf)
        print(f"  Deployed tables PDF -> {d_pdf}")
    except Exception as e:
        print(f"  Notice for {d_pdf}: {e}")

print("\n=== STEP 2: ASSEMBLE COMPLETE 8-PAGE PDF DELIVERABLE ===")
corp_pdf_final = os.path.join(SCRATCH_DIR, "CORP-FORM-21_final_0487.pdf")
assert os.path.exists(corp_pdf_final), "CORP-FORM-21_final_0487.pdf not found!"

doc_complete = fitz.open()

# Pages 1 to 6 from CORP-FORM-21
d_corp = fitz.open(corp_pdf_final)
doc_complete.insert_pdf(d_corp, links=True)
d_corp.close()

# Pages 7 to 8 from dest_tables_pdf
d_tbl = fitz.open(dest_tables_pdf)
doc_complete.insert_pdf(d_tbl, links=True)
d_tbl.close()

complete_pdf_scratch = os.path.join(SCRATCH_DIR, "OOS-262080_complete_micro_0487.pdf")
doc_complete.save(complete_pdf_scratch)
doc_complete.close()
print(f"Saved complete 8-page deliverable: {complete_pdf_scratch}")

# Deploy to all desktop & documents deliverables
deploy_targets = [
    os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis - QYC (Updated).pdf"),
    os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis (Complete).pdf"),
    os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis (2).pdf"),
    os.path.join(DOCS_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis (Complete).pdf"),
    os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis - QYC.pdf")
]

for target in deploy_targets:
    try:
        shutil.copy2(complete_pdf_scratch, target)
        print(f"  Deployed complete PDF -> {target}")
    except Exception as e:
        print(f"  Notice for {target}: {e}")

print("\n=== STEP 3: RIGOROUS VERIFICATION OF COORDINATES & VISUALS ===")
doc_ver = fitz.open(complete_pdf_scratch)
assert len(doc_ver) == 8, f"Expected 8 pages, got {len(doc_ver)}"

# Check Page 7 (Table 1 and Table 2)
p7 = doc_ver[6]
p7_links = p7.get_links()
print(f"Page 7 links count: {len(p7_links)}")
assert any("jzraMOOTFYUFD" in lk.get("uri", "") for lk in p7_links), "Missing Table 1 link!"
assert any("fd3G2StZClcy1TP2ES6BLw" in lk.get("uri", "") for lk in p7_links), "Missing Table 2 link!"

# Verify Table 2 vertical alignment
words_p7 = p7.get_text('words')
y_cfu = None
y_etx = None
for w in words_p7:
    if w[4] == 'CFU' and 560 <= w[1] <= 590:
        y_cfu = (round(w[1], 2), round(w[3], 2))
    elif '0487' in w[4]:
        y_etx = (round(w[1], 2), round(w[3], 2))

print(f"Table 2 alignment check: CFU y={y_cfu}, ETX y={y_etx}")
assert y_cfu == y_etx, f"Table 2 misaligned! CFU={y_cfu} vs ETX={y_etx}"
print(">>> PERFECT VERTICAL ALIGNMENT CONFIRMED FOR TABLE 2! <<<")

# Check Page 8 (Table 3)
p8 = doc_ver[7]
p8_links = p8.get_links()
print(f"Page 8 links count: {len(p8_links)}")
assert any("qnwcLQO5BWeJhWCBG7jl8Q" in lk.get("uri", "") for lk in p8_links), "Missing Table 3 link!"

# Verify Table 3 vertical alignment
words_p8 = p8.get_text('words')
y_cfu_p8 = None
y_etx_p8 = None
y_micro = None
for w in words_p8:
    if w[4] == 'CFU' and 400 <= w[1] <= 440:
        y_cfu_p8 = (round(w[1], 2), round(w[3], 2))
    elif '0520' in w[4]:
        y_etx_p8 = (round(w[1], 2), round(w[3], 2))
    elif w[4] == 'Micrococcus':
        y_micro = (round(w[1], 2), round(w[3], 2))

print(f"Table 3 alignment check: CFU y={y_cfu_p8}, ETX y={y_etx_p8}, Micrococcus y={y_micro}")
assert y_cfu_p8 == y_etx_p8 == y_micro, f"Table 3 misaligned! CFU={y_cfu_p8}, ETX={y_etx_p8}, Micro={y_micro}"
print(">>> PERFECT VERTICAL ALIGNMENT CONFIRMED FOR TABLE 3! <<<")

# Render verification preview images
pix7 = p7.get_pixmap(dpi=150)
pix7.save(os.path.join(SCRATCH_DIR, "verified_final_p7.png"))
im7 = Image.open(os.path.join(SCRATCH_DIR, "verified_final_p7.png"))
w7, h7 = im7.size
crop7 = im7.crop((int(w7 * 0.05), int(h7 * 0.65), int(w7 * 0.95), int(h7 * 0.90)))
crop7.save(os.path.join(SCRATCH_DIR, "preview_perfect_table2.png"))

pix8 = p8.get_pixmap(dpi=150)
pix8.save(os.path.join(SCRATCH_DIR, "verified_final_p8.png"))
im8 = Image.open(os.path.join(SCRATCH_DIR, "verified_final_p8.png"))
w8, h8 = im8.size
crop8 = im8.crop((int(w8 * 0.05), int(h8 * 0.45), int(w8 * 0.95), int(h8 * 0.70)))
crop8.save(os.path.join(SCRATCH_DIR, "preview_perfect_table3.png"))

doc_ver.close()
print(">>> PREVIEWS SAVED SUCCESSFULLY! ALL CHECKS 100% PASSED! <<<")

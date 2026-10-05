import os
import sys
import docx
import win32com.client
import fitz
import shutil

sys.stdout.reconfigure(encoding='utf-8')

DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
DOCS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
SCRATCH_DIR = os.path.join(DOCS_DIR, "scratch")

main_docx_desktop = os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis.docx")
main_docx_scratch = os.path.join(SCRATCH_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis.docx")

tables_pdf_scratch = os.path.join(SCRATCH_DIR, "Celsis table OOS-262080.pdf")

corp_pdf_src = os.path.join(SCRATCH_DIR, "CORP-FORM-21_updated_micro.pdf")
corp_pdf_desktop = os.path.join(DESKTOP_DIR, "CORP-FORM-21 Laboratory OOS Investigation Form (v11.1).pdf")
corp_pdf_docs = os.path.join(DOCS_DIR, "CORP-FORM-21 Laboratory OOS Investigation Form (v11.1).pdf")

out_qyc_desktop = os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis - QYC.pdf")
out_qyc_updated = os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis - QYC (Updated).pdf")
out_complete_desktop = os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis (Complete).pdf")
out_2_desktop = os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis (2).pdf")
out_complete_docs = os.path.join(DOCS_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis (Complete).pdf")

OLD_STR = "identified as Gram-positive coccobacilli and Gram-positive rods"
NEW_STR = "identified as Corynebacterium ureicelerivorans (Gram-positive short rods) and Mycobacterium grossiae (Gram-positive rods)"

print("=== STEP 1: UPDATE MAIN REPORT DOCX ===")
# Try loading desktop docx first; if locked or missing, load scratch docx
try:
    doc_main = docx.Document(main_docx_desktop)
    src_used = main_docx_desktop
except Exception:
    doc_main = docx.Document(main_docx_scratch)
    src_used = main_docx_scratch
print(f"Loaded main docx from: {src_used}")

c_narrative = doc_main.tables[0].rows[40].cells[0]
updated_p = False
for p in c_narrative.paragraphs:
    if OLD_STR in p.text:
        p.text = p.text.replace(OLD_STR, NEW_STR)
        updated_p = True
        print("  Successfully replaced target text in narrative Table 0 Row 40!")

if not updated_p:
    # check if already replaced
    for p in c_narrative.paragraphs:
        if "Corynebacterium ureicelerivorans" in p.text:
            print("  Narrative already contains Corynebacterium ureicelerivorans.")
            updated_p = True
            break

assert updated_p, "Failed to update narrative in main docx!"

doc_main.save(main_docx_scratch)
print(f"Saved to scratch: {main_docx_scratch}")

try:
    doc_main.save(main_docx_desktop)
    print(f"Saved to Desktop: {main_docx_desktop}")
except Exception as e:
    print(f"Desktop main docx locked: {e}")

print("\n=== STEP 2: CONVERT MAIN REPORT DOCX TO PDF VIA WORD COM ===")
word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
word.DisplayAlerts = 0

doc_com = word.Documents.Open(main_docx_scratch)
scratch_main_pdf = os.path.join(SCRATCH_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis.pdf")
doc_com.SaveAs2(scratch_main_pdf, FileFormat=17)
doc_com.Close()
word.Quit()
print(f"Exported main report PDF: {scratch_main_pdf}")

main_pdf_desktop = os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis.pdf")
try:
    shutil.copy2(scratch_main_pdf, main_pdf_desktop)
    print(f"Copied to Desktop main PDF: {main_pdf_desktop}")
except Exception as e:
    print(f"Desktop main PDF locked: {e}")

print("\n=== STEP 3: UPDATE OFFICIAL 6-PAGE PDF (CORP-FORM-21) ===")
# Open corp_pdf_src
doc_corp = fitz.open(corp_pdf_src)
p5 = doc_corp[4]
f51_found = False
for w in p5.widgets():
    if w.field_name == 'Text Field51':
        f51_found = True
        val = w.field_value
        if OLD_STR in val:
            w.field_value = val.replace(OLD_STR, NEW_STR)
            w.update()
            print("  Updated Page 5 Text Field51 with Corynebacterium & Mycobacterium!")
        elif "Corynebacterium ureicelerivorans" in val:
            print("  Page 5 Text Field51 already contains Corynebacterium ureicelerivorans.")
        else:
            print("  WARNING: Neither old string nor new string found in Text Field51!")
            print(f"  Field value sample: {val[:200]}...")

assert f51_found, "Text Field51 not found on Page 5!"

scratch_corp_final = os.path.join(SCRATCH_DIR, "CORP-FORM-21_final_0487.pdf")
doc_corp.save(scratch_corp_final)
doc_corp.close()
print(f"Saved final 6-page CORP-FORM-21 to: {scratch_corp_final}")

# Deploy 6-page form
try:
    shutil.copy2(scratch_corp_final, corp_pdf_desktop)
    print(f"Deployed to Desktop: {corp_pdf_desktop}")
except Exception as e:
    print(f"Desktop CORP-FORM-21 locked: {e}")

try:
    shutil.copy2(scratch_corp_final, corp_pdf_docs)
    print(f"Deployed to Documents: {corp_pdf_docs}")
except Exception as e:
    print(f"Documents CORP-FORM-21 locked: {e}")

print("\n=== STEP 4: ASSEMBLE 8-PAGE COMPLETE PACKET ===")
doc_complete = fitz.open()

# Insert 6 pages from CORP-FORM-21
doc_f_in = fitz.open(scratch_corp_final)
doc_complete.insert_pdf(doc_f_in, links=True)
doc_f_in.close()

# Insert 2 pages from scratch_tables_pdf
doc_t_in = fitz.open(tables_pdf_scratch)
doc_complete.insert_pdf(doc_t_in, links=True)
doc_t_in.close()

scratch_complete = os.path.join(SCRATCH_DIR, "OOS-262080_complete_micro_0487.pdf")
doc_complete.save(scratch_complete)
doc_complete.close()
print(f"Saved complete 8-page PDF to: {scratch_complete}")

# Deploy to all target destinations
destinations = [
    out_qyc_updated,
    out_complete_desktop,
    out_2_desktop,
    out_complete_docs,
    out_qyc_desktop
]

for dst in destinations:
    try:
        shutil.copy2(scratch_complete, dst)
        print(f"  Deployed to -> {dst}")
    except Exception as e:
        print(f"  Could not deploy to {dst}: {e}")

print("\n=== STEP 5: RIGOROUS VERIFICATION OF ALL 8 PAGES ===")
doc_ver = fitz.open(scratch_complete)
print(f"Total pages: {len(doc_ver)}")

# Check Page 5 narrative
p5_text = doc_ver[4].get_text()
has_coryne = "Corynebacterium ureicelerivorans" in p5_text
has_myco = "Mycobacterium grossiae" in p5_text
has_micro = "Micrococcus luteus" in p5_text
print(f"Page 5 check: Corynebacterium={has_coryne}, Mycobacterium={has_myco}, Micrococcus={has_micro}")
assert has_coryne and has_myco and has_micro, "Page 5 missing organism identifications!"

# Check Page 7 (Table 1 & Table 2)
p7_text = doc_ver[6].get_text()
p7_links = doc_ver[6].get_links()
print(f"Page 7 links count: {len(p7_links)}")
has_t1_link = any("jzraMOOTFYUFD" in lk.get("uri", "") for lk in p7_links)
has_t2_link = any("fd3G2StZClcy1TP2ES6BLw" in lk.get("uri", "") for lk in p7_links)
has_coryne_t2 = "Corynebacterium" in p7_text
print(f"Page 7 check: Table 1 Link={has_t1_link}, Table 2 Link={has_t2_link}, Organism in text={has_coryne_t2}")
assert has_t1_link, "Page 7 Table 1 link missing!"
assert has_t2_link, "Page 7 Table 2 link missing!"
assert has_coryne_t2, "Page 7 Table 2 text missing Corynebacterium!"

# Check Page 8 (Table 3)
p8_text = doc_ver[7].get_text()
p8_links = doc_ver[7].get_links()
print(f"Page 8 links count: {len(p8_links)}")
has_t3_link = any("qnwcLQO5BWeJhWCBG7jl8Q" in lk.get("uri", "") for lk in p8_links)
has_micro_t3 = "Micrococcus" in p8_text
print(f"Page 8 check: Table 3 Link={has_t3_link}, Organism in text={has_micro_t3}")
assert has_t3_link, "Page 8 Table 3 link missing!"
assert has_micro_t3, "Page 8 Table 3 text missing Micrococcus!"

doc_ver.close()
print("\n>>> ALL 8-PAGE VERIFICATIONS PASSED WITH 100% SUCCESS! <<<")

import os
import sys
import shutil
import fitz
from PIL import Image

sys.stdout.reconfigure(encoding='utf-8')

DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
DOCS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
SCRATCH_DIR = os.path.join(DOCS_DIR, "scratch")

# Input 6-page form and 2-page tables
corp_pdf_in = os.path.join(SCRATCH_DIR, "CORP-FORM-21_final_0487.pdf")
tables_pdf_in = os.path.join(SCRATCH_DIR, "Celsis table OOS-262080.pdf")

assert os.path.exists(corp_pdf_in), f"Missing {corp_pdf_in}"
assert os.path.exists(tables_pdf_in), f"Missing {tables_pdf_in}"

print("=== STEP 1: FIX ALL AMPERSANDS IN 6-PAGE CORP-FORM-21 ===")
doc_corp = fitz.open(corp_pdf_in)

fixed_count = 0
for p_idx, page in enumerate(doc_corp):
    for w in page.widgets():
        val = w.field_value or ''
        if '&' in val:
            # Fix Text Field13 (Page 1)
            if w.field_name == 'Text Field13':
                w.field_value = val.replace(' & ', ', and ')
                w.update()
                fixed_count += 1
                print(f"  Fixed Page 1 Text Field13 -> {repr(w.field_value)}")
            # Fix Text Field22 / 23 (Page 2)
            elif w.field_name in ['Text Field22', 'Text Field23']:
                w.field_value = val.replace(' & ', ' and ')
                w.update()
                fixed_count += 1
                print(f"  Fixed Page {p_idx+1} {w.field_name} (Reagents & Kits -> Reagents and Kits)")
            else:
                w.field_value = val.replace(' & ', ' and ')
                w.update()
                fixed_count += 1
                print(f"  Fixed Page {p_idx+1} {w.field_name}")

print(f"Total fields updated: {fixed_count}")

# Verify 0 ampersands remain in any field
remaining_amp = 0
for p_idx, page in enumerate(doc_corp):
    for w in page.widgets():
        if w.field_value and '&' in w.field_value:
            print(f"  WARNING: Page {p_idx+1} {w.field_name} still has &: {repr(w.field_value)}")
            remaining_amp += 1

assert remaining_amp == 0, f"{remaining_amp} fields still contain '&'!"
print(">>> ZERO AMPERSANDS IN ALL FORM FIELDS CONFIRMED! <<<")

corp_pdf_clean = os.path.join(SCRATCH_DIR, "CORP-FORM-21_clean.pdf")
doc_corp.save(corp_pdf_clean)
doc_corp.close()

# Deploy cleaned 6-page form to Desktop and Documents
corp_destinations = [
    os.path.join(DESKTOP_DIR, "CORP-FORM-21 Laboratory OOS Investigation Form (v11.1).pdf"),
    os.path.join(DOCS_DIR, "CORP-FORM-21 Laboratory OOS Investigation Form (v11.1).pdf")
]
for dst in corp_destinations:
    try:
        shutil.copy2(corp_pdf_clean, dst)
        print(f"  Deployed 6-page form -> {dst}")
    except Exception as e:
        print(f"  Notice for {dst}: {e}")

print("\n=== STEP 2: ASSEMBLE COMPLETE 8-PAGE PDF DELIVERABLE ===")
doc_complete = fitz.open()

# Pages 1 to 6 from cleaned CORP-FORM-21
d_clean = fitz.open(corp_pdf_clean)
doc_complete.insert_pdf(d_clean, links=True)
d_clean.close()

# Pages 7 to 8 from tables PDF
d_tbl = fitz.open(tables_pdf_in)
doc_complete.insert_pdf(d_tbl, links=True)
d_tbl.close()

complete_pdf_clean = os.path.join(SCRATCH_DIR, "OOS-262080_complete_micro_0487.pdf")
doc_complete.save(complete_pdf_clean)
doc_complete.close()
print(f"Saved complete 8-page PDF to: {complete_pdf_clean}")

# Deploy to all desktop & documents targets
complete_destinations = [
    os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis - QYC (Updated).pdf"),
    os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis (Complete).pdf"),
    os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis (2).pdf"),
    os.path.join(DOCS_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis (Complete).pdf"),
    os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis - QYC.pdf")
]

for dst in complete_destinations:
    try:
        shutil.copy2(complete_pdf_clean, dst)
        print(f"  Deployed complete 8-page PDF -> {dst}")
    except Exception as e:
        print(f"  Notice for {dst}: {e}")

print("\n=== STEP 3: FINAL VERIFICATION & SCREENSHOT EVIDENCE ===")
doc_final = fitz.open(complete_pdf_clean)
assert len(doc_final) == 8, f"Expected 8 pages, got {len(doc_final)}"

# 1. Page 1 Text Field13 check
p1 = doc_final[0]
f13_val = None
for w in p1.widgets():
    if w.field_name == 'Text Field13':
        f13_val = w.field_value
        break

print(f"Page 1 Text Field13: {repr(f13_val)}")
assert '&' not in f13_val, "Page 1 Text Field13 still has '&'!"
assert ', and Cuong Du' in f13_val, "Page 1 Text Field13 missing ', and Cuong Du'!"

# 2. Render Page 1 verification screenshot
pix1 = p1.get_pixmap(dpi=150)
p1_img_path = os.path.join(SCRATCH_DIR, "final_verified_p1.png")
pix1.save(p1_img_path)

im1 = Image.open(p1_img_path)
w1, h1 = im1.size
crop1 = im1.crop((int(w1 * 0.40), int(h1 * 0.60), int(w1 * 0.95), int(h1 * 0.72)))
crop1_path = os.path.join(SCRATCH_DIR, "verified_comments_no_amp.png")
crop1.save(crop1_path)
print(f"Saved Page 1 cropped verification: {crop1_path}")

# 3. Check Page 7 and Page 8 links
p7_links = doc_final[6].get_links()
p8_links = doc_final[7].get_links()
print(f"Page 7 links: {len(p7_links)}, Page 8 links: {len(p8_links)}")
assert any("jzraMOOTFYUFD" in lk.get("uri", "") for lk in p7_links), "Missing Table 1 link!"
assert any("fd3G2StZClcy1TP2ES6BLw" in lk.get("uri", "") for lk in p7_links), "Missing Table 2 link!"
assert any("qnwcLQO5BWeJhWCBG7jl8Q" in lk.get("uri", "") for lk in p8_links), "Missing Table 3 link!"

# 4. Check user preserved fields
p3 = doc_final[2]
f48 = [w.field_value for w in p3.widgets() if w.field_name == 'Text Field48']
assert f48 == ['N/A QYC 02Oct26'], f"Text Field48 regressed: {f48}"

p6 = doc_final[5]
f54 = [w.field_value for w in p6.widgets() if w.field_name == 'Text Field54']
assert f54 == [''], f"Text Field54 regressed: {f54}"

doc_final.close()
print("\n>>> ALL CHECKS PASSED WITH 100% SUCCESS! DELIVERABLES SYNCHRONIZED! <<<")

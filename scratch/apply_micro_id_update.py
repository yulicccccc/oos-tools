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

print("=== STEP 1: UPDATE MAIN REPORT DOCX ===")
doc_main = docx.Document(main_docx_path)

# Update narrative inside Table 0 Row 40
c_narrative = doc_main.tables[0].rows[40].cells[0]
p_updated = False
for p in c_narrative.paragraphs:
    if "identified as Gram-positive cocci" in p.text:
        p.text = p.text.replace("identified as Gram-positive cocci", "identified as Micrococcus luteus (Gram-positive cocci)")
        p_updated = True
        print("  Updated narrative in Table 0 Row 40 cell.")

assert p_updated, "Could not find target narrative text in Table 0 Row 40!"

# Update Table in main docx (Table 2 or Table 3)
t_updated = False
for t_idx, t in enumerate(doc_main.tables[1:], 1):
    for r in t.rows:
        text = ' '.join(c.text for c in r.cells)
        if SAMPLE_0520 in text:
            for c in r.cells:
                if 'Gram (+) cocci' in c.text:
                    for p in c.paragraphs:
                        p.text = ""
                    run = c.paragraphs[0].add_run("Micrococcus luteus")
                    run.font.name = "Times New Roman"
                    run.font.size = docx.shared.Pt(7)
                    c.paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
                    t_updated = True
                    print(f"  Updated Microbial ID to 'Micrococcus luteus' in main docx Table {t_idx}.")
            # Add hyperlink to ETX ID cell
            for c in r.cells:
                if SAMPLE_0520 in c.text:
                    part = c.part
                    r_id = part.relate_to(URL_0520, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
                    tc_xml = (
                        f'<w:tc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                        f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                        f'<w:tcPr>'
                        f'<w:tcW w:w="1444" w:type="dxa"/>'
                        f'<w:tcBorders>'
                        f'<w:top w:val="single" w:sz="8" w:space="0" w:color="auto"/>'
                        f'<w:left w:val="single" w:sz="8" w:space="0" w:color="auto"/>'
                        f'<w:bottom w:val="single" w:sz="8" w:space="0" w:color="auto"/>'
                        f'<w:right w:val="single" w:sz="8" w:space="0" w:color="auto"/>'
                        f'</w:tcBorders>'
                        f'<w:vAlign w:val="center"/>'
                        f'</w:tcPr>'
                        f'<w:p>'
                        f'<w:pPr>'
                        f'<w:spacing w:line="360" w:lineRule="auto"/>'
                        f'<w:jc w:val="center"/>'
                        f'<w:rPr>'
                        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
                        f'<w:sz w:val="14"/>'
                        f'<w:szCs w:val="14"/>'
                        f'</w:rPr>'
                        f'</w:pPr>'
                        f'<w:hyperlink r:id="{r_id}" w:history="1">'
                        f'<w:r>'
                        f'<w:rPr>'
                        f'<w:rStyle w:val="Hyperlink"/>'
                        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
                        f'<w:color w:val="0000FF"/>'
                        f'<w:sz w:val="14"/>'
                        f'<w:szCs w:val="14"/>'
                        f'<w:u w:val="single"/>'
                        f'</w:rPr>'
                        f'<w:t>{SAMPLE_0520}</w:t>'
                        f'</w:r>'
                        f'</w:hyperlink>'
                        f'</w:p>'
                        f'</w:tc>'
                    )
                    new_tc = parse_xml(tc_xml)
                    c._tc.getparent().replace(c._tc, new_tc)
                    print(f"  Inserted hyperlink for {SAMPLE_0520} in main docx Table {t_idx}.")

doc_main.save(main_docx_path)
print(f"Saved updated main DOCX: {main_docx_path}")

print("\n=== STEP 2: CONVERT MAIN REPORT DOCX TO PDF VIA WORD COM ===")
word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
word.DisplayAlerts = 0

doc_com = word.Documents.Open(main_docx_path)
main_pdf_path = os.path.join(DESKTOP_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis.pdf")
scratch_main_pdf = os.path.join(SCRATCH_DIR, "OOS-262080 Optimal Balance Pharmacy (E19193) - Celsis.pdf")
doc_com.SaveAs2(scratch_main_pdf, FileFormat=17)
doc_com.Close()
word.Quit()
shutil.copy2(scratch_main_pdf, main_pdf_path)
print(f"Converted and deployed main PDF to: {main_pdf_path}")

print("\n=== STEP 3: UPDATE OFFICIAL 6-PAGE PDF (CORP-FORM-21) ===")
# We load CORP-FORM-21, update only the narrative on Page 5
doc_form = fitz.open(out_corp_desktop)
p5 = doc_form[4]
for w in p5.widgets():
    if w.field_name == 'Text Field51':
        w.field_value = w.field_value.replace(
            "identified as Gram-positive cocci",
            "identified as Micrococcus luteus (Gram-positive cocci)"
        )
        w.update()
        print("  Updated Page 5 Text Field51 with Micrococcus luteus.")

scratch_corp_updated = os.path.join(SCRATCH_DIR, "CORP-FORM-21_updated_micro.pdf")
doc_form.save(scratch_corp_updated)
doc_form.close()

# Deploy updated 6-page CORP-FORM-21
shutil.copy2(scratch_corp_updated, out_corp_desktop)
shutil.copy2(scratch_corp_updated, out_corp_docs)
print(f"Deployed updated CORP-FORM-21 to Desktop and Documents.")

print("\n=== STEP 4: ASSEMBLE 8-PAGE COMPLETE PACKET WITH UPDATED TABLES ===")
doc_complete = fitz.open()
doc_form_in = fitz.open(scratch_corp_updated)
doc_complete.insert_pdf(doc_form_in, links=True)
doc_form_in.close()

doc_tbl_in = fitz.open(tables_pdf_path)
doc_complete.insert_pdf(doc_tbl_in, links=True)
doc_tbl_in.close()

scratch_complete_updated = os.path.join(SCRATCH_DIR, "OOS-262080_complete_micro.pdf")
doc_complete.save(scratch_complete_updated)
doc_complete.close()

# Deploy 8-page complete deliverables
for dst in [out_qyc_desktop, out_complete_desktop, out_2_desktop, out_complete_docs]:
    shutil.copy2(scratch_complete_updated, dst)
    print(f"  Deployed complete 8-page PDF -> {dst}")

print("\n=== STEP 5: VERIFY COMPLETE PDF INTEGRITY ===")
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

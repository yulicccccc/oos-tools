import os, sys, shutil, datetime
import docx
from docx.shared import Pt, Inches
from docx.oxml import parse_xml
from docx.opc.constants import RELATIONSHIP_TYPE
from docxtpl import DocxTemplate
import win32com.client
from pypdf import PdfReader
import fitz

OOS_ROOT = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
sys.path.insert(0, OOS_ROOT)
os.chdir(OOS_ROOT)

OUTPUT_DIR = os.path.join(OOS_ROOT, "scratch")
DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
DOCUMENTS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents"

# Base template from proven pure Celsis style without aliquoting
TPL_BASE = os.path.join(OUTPUT_DIR, "tables_71_pure_celsis.docx")
if not os.path.exists(TPL_BASE):
    TPL_BASE = os.path.join(OOS_ROOT, "tables for celsis.docx")

data = {
    "oos_id": "262098",
    "client_name": "Solyn LLC (E75000)",
    "sample_id": "ETX-260902-0505",
    "sample_url": "https://etrax.eagleanalytical.com/Submission/Details/%247Ydn%24JOiYgmsnaKjbxn-g__",
    "sample_name": "GLP3R/Cagrilinitide",
    "lot_number": "2608-216",
    "analyst_name": "Alex Saravia",
    "reading_name": "Elysse Nioupin",
    "positive_id": "ETX-260910-0290",
    "positive_url": "https://etrax.eagleanalytical.com/Submission/Details?id=oZa2TD1UI-0eAJOFysBJ3A__",
    "positive_media": "1 x 100mL TSB bottle",
    "positive_org": "Pending\n(Gram (-) rods)",

    # Table 2 Dates & Analysts
    "process_date": "04Sep26",
    "pro_before_test": "03Sep26",
    "pro_test_date": "04Sep26",
    "pro_after_test": "08Sep26",
    "pro_analyst_initial": "ES",
    "pro_date_of_weekly": "04Sep26",

    # Personnel EM
    "pro_be_obs_pers_dur_pro": "No growth", "pro_be_etx_pers_dur_pro": "N/A", "pro_be_id_pers_dur_pro": "N/A",
    "pro_obs_pers_dur_pro": "No growth",    "pro_etx_pers_dur_pro": "N/A", "pro_id_pers_dur_pro": "N/A",
    "pro_af_obs_pers_dur_pro": "No growth", "pro_af_etx_pers_dur_pro": "N/A", "pro_af_id_pers_dur_pro": "N/A",

    # Surface ISO 5
    "pro_be_obs_surf_dur_pro": "No growth", "pro_be_etx_surf_dur_pro": "N/A", "pro_be_id_surf_dur_pro": "N/A",
    "pro_obs_surf_dur_pro": "No growth",    "pro_etx_surf_dur_pro": "N/A", "pro_id_surf_dur_pro": "N/A",
    "pro_af_obs_surf_dur_pro": "No growth", "pro_af_etx_surf_dur_pro": "N/A", "pro_af_id_surf_dur_pro": "N/A",

    # Settling ISO 5
    "pro_be_obs_sett_dur_pro": "No growth", "pro_be_etx_sett_dur_pro": "N/A", "pro_be_id_sett_dur_pro": "N/A",
    "pro_obs_sett_dur_pro": "No growth",    "pro_etx_sett_dur_pro": "N/A", "pro_id_sett_dur_pro": "N/A",
    "pro_af_obs_sett_dur_pro": "No growth", "pro_af_etx_sett_dur_pro": "N/A", "pro_af_id_sett_dur_pro": "N/A",

    # Weekly Active Air & Surface
    "pro_obs_air_wk_of": "1 CFU (ISO 8 114)",
    "pro_etx_air_wk_of": "ETX-260914-0487",
    "pro_id_air_wk_of": "Pending",
    "pro_obs_room_wk_of": "No growth",
    "pro_etx_room_wk_of": "N/A",
    "pro_id_room_wk_of": "N/A"
}

OUT_DOCX = os.path.join(OUTPUT_DIR, f"Tables OOS-{data['oos_id']} {data['client_name']} - USP71.docx")
OUT_PDF = os.path.join(OUTPUT_DIR, f"Tables OOS-{data['oos_id']} {data['client_name']} - USP71.pdf")

tpl = DocxTemplate(TPL_BASE)
tpl.render(data)
tpl.save(OUT_DOCX)

# Post-processing: Add Native Hyperlinks in Table 1 with strict Times New Roman
doc_rendered = docx.Document(OUT_DOCX)

def add_hyperlink_to_cell(cell, url, text):
    tc = cell._tc
    for p in tc.xpath('w:p'):
        tc.remove(p)
    part = cell.part
    r_id = part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    p_xml = (
        f'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<w:pPr>'
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
        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
        f'<w:color w:val="0000FF"/>'
        f'<w:sz w:val="14"/>'
        f'<w:szCs w:val="14"/>'
        f'<w:u w:val="single"/>'
        f'</w:rPr>'
        f'<w:t>{text}</w:t>'
        f'</w:r>'
        f'</w:hyperlink>'
        f'</w:p>'
    )
    tc.append(parse_xml(p_xml))
    cell.vertical_alignment = docx.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER

t1_rendered = doc_rendered.tables[0]
add_hyperlink_to_cell(t1_rendered.rows[1].cells[2], data["sample_url"], data["sample_id"])
add_hyperlink_to_cell(t1_rendered.rows[1].cells[3], data["positive_url"], data["positive_id"])

# Strict 100% Times New Roman Enforcement across all runs in all tables
for t_idx, table in enumerate(doc_rendered.tables):
    for r_idx, row in enumerate(table.rows):
        for c_idx, cell in enumerate(row.cells):
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.name = "Times New Roman"
                    rPr = run._r.get_or_add_rPr()
                    rFonts = parse_xml('<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>')
                    for existing in rPr.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts'):
                        rPr.remove(existing)
                    rPr.append(rFonts)

# Row heights
for t in doc_rendered.tables:
    for r in t.rows:
        trPr = r._tr.get_or_add_trPr()
        trHeight = parse_xml(f'<w:trHeight xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="220" w:hRule="atLeast"/>')
        trPr.append(trHeight)

doc_rendered.save(OUT_DOCX)
print("Saved DOCX:", OUT_DOCX)

# Convert to PDF via Word COM
try:
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try:
        d = word.Documents.Open(os.path.abspath(OUT_DOCX), ReadOnly=True)
        d.SaveAs(os.path.abspath(OUT_PDF), FileFormat=17)
        d.Close(SaveChanges=False)
        print("Exported PDF via Word COM:", OUT_PDF)
    finally:
        word.Quit()
except Exception as e:
    print(f"Word COM error: {e}")

if os.path.exists(OUT_PDF):
    reader = PdfReader(OUT_PDF)
    print(f"Verified PDF Page Count: {len(reader.pages)} (Must be 1)")

# Sync to Desktop and Documents safely
for dest_dir in [DESKTOP_DIR, DOCUMENTS_DIR]:
    dest_docx = os.path.join(dest_dir, os.path.basename(OUT_DOCX))
    dest_pdf = os.path.join(dest_dir, os.path.basename(OUT_PDF))
    try:
        shutil.copy2(OUT_DOCX, dest_docx)
        print("Synced DOCX to:", dest_docx)
    except Exception as e:
        print(f"Notice: {dest_docx} copy skipped ({e})")
    if os.path.exists(OUT_PDF):
        try:
            shutil.copy2(OUT_PDF, dest_pdf)
            print("Synced PDF to:", dest_pdf)
        except Exception as e:
            print(f"Notice: {dest_pdf} copy skipped ({e})")

print("All standalone table files generated successfully.")

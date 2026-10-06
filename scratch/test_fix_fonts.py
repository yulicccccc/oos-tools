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
CELSIS_TPL = os.path.join(OOS_ROOT, "tables for celsis.docx")

# 1. Load original tables for celsis.docx
doc_base = docx.Document(CELSIS_TPL)

# In Table 1: update Aliquoting Analyst to Reading Analyst preserving exact run fonts
t1 = doc_base.tables[0]

# Row 0, Cell 1 (Header)
cell_h = t1.rows[0].cells[1]
cell_h.text = ''
p_h = cell_h.paragraphs[0]
p_h.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
r_h = p_h.add_run('Reading Analyst')
r_h.font.name = 'Times New Roman'
r_h.font.size = Pt(7.0)
r_h.bold = True
rPr_h = r_h._r.get_or_add_rPr()
rFonts_h = parse_xml('<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>')
rPr_h.append(rFonts_h)

# Row 1, Cell 1 (Data tag)
cell_d = t1.rows[1].cells[1]
cell_d.text = ''
p_d = cell_d.paragraphs[0]
p_d.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
r_d = p_d.add_run('{{ reading_name }}')
r_d.font.name = 'Times New Roman'
r_d.font.size = Pt(7.0)
r_d.bold = False
rPr_d = r_d._r.get_or_add_rPr()
rFonts_d = parse_xml('<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>')
rPr_d.append(rFonts_d)

# 2. In Table 2: LEAVE ALL HEADER ROWS (Row 1, 5, 12, 14) 100% UNTOUCHED!
# They are original Celsis formatting: Times New Roman, bold, 7.0pt.

# 3. Remove Table 3 (Aliquoting Table) completely
t3 = doc_base.tables[2]
t3._element.getparent().remove(t3._element)

# Remove Table 3 heading paragraph
for p in list(doc_base.paragraphs):
    if "Table 3:" in p.text or "Aliquoting" in p.text:
        p._element.getparent().remove(p._element)

TPL_CLEAN = os.path.join(OUTPUT_DIR, "tables_71_pure_celsis.docx")
doc_base.save(TPL_CLEAN)
print("Saved clean template to:", TPL_CLEAN)

# 4. Context for OOS-262064
context = {
    # Table 1
    "sample_id": "ETX-260804-0335",
    "analyst_name": "Pending",
    "reading_name": "Qiyue Chen",
    "positive_id": "ETX-260904-0331",
    "positive_media": "1 x 100mL TSB bottle",
    "positive_org": "Pending",

    # Table 2
    "process_date": "25Aug26",
    "pro_before_test": "24Aug26",
    "pro_test_date": "25Aug26",
    "pro_after_test": "26Aug26",
    "pro_analyst_initial": "Pending",
    "pro_date_of_weekly": "25Aug26",

    # Personnel
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
    "pro_af_obs_sett_dur_pro": "No growth", "pro_af_etx_sett_dur_pro": "N/A", "pro_id_sett_dur_pro": "N/A",

    # Weekly Active Air & Surface
    "pro_obs_air_wk_of": "4 CFU (ISO 8 114)",
    "pro_etx_air_wk_of": "ETX-260901-0112",
    "pro_id_air_wk_of": "3 Gram (+) cocci,\n1 Hyphae",
    "pro_obs_room_wk_of": "No growth",
    "pro_etx_room_wk_of": "N/A",
    "pro_id_room_wk_of": "N/A"
}

OUT_DOCX = os.path.join(OUTPUT_DIR, "Tables OOS-262064 Blue Ocean Rx, LLC (E75110) - USP71.docx")
OUT_PDF = os.path.join(OUTPUT_DIR, "Tables OOS-262064 Blue Ocean Rx, LLC (E75110) - USP71.pdf")

tpl = DocxTemplate(TPL_CLEAN)
tpl.render(context)
tpl.save(OUT_DOCX)

# 5. Post-process: add native hyperlinks in Table 1 with strict Times New Roman
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
        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
        f'<w:sz w:val="14"/>'
        f'<w:szCs w:val="14"/>'
        f'</w:rPr>'
        f'</w:pPr>'
        f'<w:hyperlink r:id="{r_id}" w:history="1">'
        f'<w:r>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
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
URL_SAMPLE = "https://etrax.eagleanalytical.com/Submission/Details/ETX-260804-0335"
URL_MICRO = "https://etrax.eagleanalytical.com/Submission/Details?id=qa0T20auUnDSY-CQtpZQpw__"
add_hyperlink_to_cell(t1_rendered.rows[1].cells[2], URL_SAMPLE, context["sample_id"])
add_hyperlink_to_cell(t1_rendered.rows[1].cells[3], URL_MICRO, context["positive_id"])

# 6. UNIVERSAL FONT ENFORCEMENT: Enforce Times New Roman on every run in every cell!
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

doc_rendered.save(OUT_DOCX)
print("Rendered and verified DOCX:", OUT_DOCX)

# 7. Convert to PDF via Word COM
try:
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try:
        d = word.Documents.Open(os.path.abspath(OUT_DOCX), ReadOnly=True)
        d.SaveAs(os.path.abspath(OUT_PDF), FileFormat=17)
        d.Close(SaveChanges=False)
        print("Exported PDF:", OUT_PDF)
    finally:
        word.Quit()
except Exception as e:
    print(f"Word COM error: {e}")

if os.path.exists(OUT_PDF):
    reader = PdfReader(OUT_PDF)
    print(f"Verified PDF Page Count: {len(reader.pages)} (Must be 1)")

# 8. Sync to Desktop and Documents
for dest_dir in [DESKTOP_DIR, DOCUMENTS_DIR]:
    shutil.copy2(OUT_DOCX, os.path.join(dest_dir, os.path.basename(OUT_DOCX)))
    if os.path.exists(OUT_PDF):
        shutil.copy2(OUT_PDF, os.path.join(dest_dir, os.path.basename(OUT_PDF)))
print("Synced files to Desktop and Documents successfully.")

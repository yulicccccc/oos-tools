import os, sys, shutil, datetime
import docx
from docx.shared import Pt, Inches
from docx.oxml import parse_xml
from docx.opc.constants import RELATIONSHIP_TYPE
from docxtpl import DocxTemplate
import win32com.client
from pypdf import PdfReader

OOS_ROOT = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS"
sys.path.insert(0, OOS_ROOT)
os.chdir(OOS_ROOT)

OUTPUT_DIR = os.path.join(OOS_ROOT, "scratch")
DESKTOP_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
DOCUMENTS_DIR = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents"
CELSIS_TPL = os.path.join(OOS_ROOT, "tables for celsis.docx")

# 1. Backup original tables for celsis.docx
backup_dir = os.path.join(OOS_ROOT, ".history")
os.makedirs(backup_dir, exist_ok=True)
timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
shutil.copy2(CELSIS_TPL, os.path.join(backup_dir, f"tables_for_celsis_backup_{timestamp}.docx"))

# 2. Build template 'tables for 71.docx' derived from Celsis template
doc_base = docx.Document(CELSIS_TPL)

# In Table 1: change Aliquoting Analyst to Reading Analyst
t1 = doc_base.tables[0]
for p in t1.rows[0].cells[1].paragraphs:
    if "Aliquoting Analyst" in p.text:
        p.text = "Reading Analyst"
for p in t1.rows[1].cells[1].paragraphs:
    if "aliquoting_name" in p.text:
        p.text = "{{ reading_name }}"

# In Table 2: BSC header Row 5
# Currently: "Biological Safety Cabinet EM Bracketing Biological Safety Cabinet (BSC)"
# Let's ensure it can display BSC ID:
t2 = doc_base.tables[1]
for c in t2.rows[5].cells:
    if "Biological Safety Cabinet (BSC)" in c.text and "{{ bsc_id }}" not in c.text:
        c.text = "Biological Safety Cabinet EM Bracketing Biological Safety Cabinet (BSC) {{ bsc_id }}"
        break

# In Table 2: Row 12 and Row 14 header text can have cleanroom name if needed
# Keep clean celsis style:
# Row 12: Weekly Active Air Sampling Bracketing {{ cr_name }}
# Row 14: Surface Sampling of Anteroom and Cleanroom Bracketing {{ cr_name }}
for c in t2.rows[12].cells:
    if "Weekly Active Air Sampling Bracketing" in c.text and "{{ cr_name }}" not in c.text:
        c.text = "Weekly Active Air Sampling of Cleanroom - {{ cr_name }}"
        break
for c in t2.rows[14].cells:
    if "Surface Sampling of Anteroom and Cleanroom Bracketing" in c.text and "{{ cr_name }}" not in c.text:
        c.text = "Surface Sampling of Anteroom and Cleanroom – {{ cr_name }}"
        break

# Remove Table 3 (Aliquoting Table)
t3 = doc_base.tables[2]
t3._element.getparent().remove(t3._element)

# Remove Table 3 heading paragraph
for p in list(doc_base.paragraphs):
    if "Table 3:" in p.text or "Aliquoting" in p.text:
        p._element.getparent().remove(p._element)

TPL_71 = os.path.join(OUTPUT_DIR, "template_tables_71_celsis_style.docx")
doc_base.save(TPL_71)
print("Saved derived 71 template:", TPL_71)

# 3. Define Context for OOS-262064
URL_SAMPLE = "https://etrax.eagleanalytical.com/Submission/Details/ETX-260804-0335"
URL_MICRO = "https://etrax.eagleanalytical.com/Submission/Details?id=qa0T20auUnDSY-CQtpZQpw__"

context = {
    # Table 1
    "sample_id": "ETX-260804-0335",
    "analyst_name": "Pending",
    "reading_name": "Qiyue Chen",
    "positive_id": "ETX-260904-0331",
    "positive_media": "1 x 100mL TSB bottle",
    "positive_org": "Pending",

    # Table 2
    "bsc_id": "E001316",
    "cr_name": "CR114 (E001736)",
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
    "pro_af_obs_sett_dur_pro": "No growth", "pro_af_etx_sett_dur_pro": "N/A", "pro_af_id_sett_dur_pro": "N/A",

    # Weekly Active Air & Surface
    "pro_obs_air_wk_of": "4 CFU (ISO 8 114)",
    "pro_etx_air_wk_of": "ETX-260901-0112",
    "pro_id_air_wk_of": "3 Gram (+) cocci,\n1 Hyphae",
    "pro_obs_room_wk_of": "No growth",
    "pro_etx_room_wk_of": "N/A",
    "pro_id_room_wk_of": "N/A"
}

# 4. Render using DocxTemplate
OUT_DOCX = os.path.join(OUTPUT_DIR, "Tables OOS-262064 Blue Ocean Rx, LLC (E75110) - USP71.docx")
OUT_PDF = os.path.join(OUTPUT_DIR, "Tables OOS-262064 Blue Ocean Rx, LLC (E75110) - USP71.pdf")

tpl = DocxTemplate(TPL_71)
tpl.render(context)
tpl.save(OUT_DOCX)

# 5. Post-Process to Add Native Hyperlinks in Table 1
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

# Format Table 1 row 0 (headers)
t1_headers = [
    "Processing Analyst",
    "Reading Analyst",
    "Sample ID",
    "Related Microbial ID",
    "Media with microbial growth",
    "Microbial ID"
]
for c_idx, h_text in enumerate(t1_headers):
    cell = t1_rendered.rows[0].cells[c_idx]
    p = cell.paragraphs[0]
    p.text = h_text
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    cell.vertical_alignment = docx.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for run in p.runs:
        run.font.name = "Times New Roman"
        run.font.size = Pt(7.5)
        run.bold = True

# Format Table 1 row 1 (data cells)
for c_idx in [0, 1, 4, 5]:
    cell = t1_rendered.rows[1].cells[c_idx]
    val = [context["analyst_name"], context["reading_name"], "", "", context["positive_media"], context["positive_org"]][c_idx]
    p = cell.paragraphs[0]
    p.text = val
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    cell.vertical_alignment = docx.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for run in p.runs:
        run.font.name = "Times New Roman"
        run.font.size = Pt(7.5)
        run.bold = False

add_hyperlink_to_cell(t1_rendered.rows[1].cells[2], URL_SAMPLE, context["sample_id"])
add_hyperlink_to_cell(t1_rendered.rows[1].cells[3], URL_MICRO, context["positive_id"])

# Also ensure tight row heights for strict 1-page fit
for t in doc_rendered.tables:
    for r in t.rows:
        trPr = r._tr.get_or_add_trPr()
        trHeight = parse_xml(f'<w:trHeight xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="220" w:hRule="atLeast"/>')
        trPr.append(trHeight)

doc_rendered.save(OUT_DOCX)
print("Rendered and post-processed DOCX:", OUT_DOCX)

# 6. Convert to PDF via Word COM
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

# 7. Sync to Desktop and Documents
for dest_dir in [DESKTOP_DIR, DOCUMENTS_DIR]:
    shutil.copy2(OUT_DOCX, os.path.join(dest_dir, os.path.basename(OUT_DOCX)))
    if os.path.exists(OUT_PDF):
        shutil.copy2(OUT_PDF, os.path.join(dest_dir, os.path.basename(OUT_PDF)))
print("Synced files to Desktop and Documents successfully.")

import os, sys, shutil
import fitz
import docx
import win32com.client as win32
from docx.shared import Inches, Pt
from docx.oxml import parse_xml

sp_04sep = [
    'Staphylococcus capitis', 'Micrococcus luteus', 'Moraxella osloensis',
    'Corynebacterium segmentosum', 'Kytococcus sedentarius', 'Staphylococcus hominis',
    'Paracoccus yeei', 'Brevibacterium sp. CS2', 'Corynebacterium ureicelerivorans',
    'Micrococcus sp.', 'Corynebacterium sanguinis', 'Staphylococcus haemolyticus',
    'Corynebacterium genitalium', 'Cordyceps memorabilis', 'Allophoma tropica',
    'Corynebacterium jeikeium', 'Pseudomonas fulva', 'Lophiotremataceae',
    'Acrodontium salmoneum', 'Quambalaria cyanescens', 'Sordariales', 'Streptococcus sp. A12'
]

sp_09sep = [
    'Corynebacterium simulans', 'Staphylococcus epidermidis', 'Staphylococcus pasteuri',
    'Micrococcus luteus', 'Staphylococcus hominis', 'Rothia dentocariosa',
    'Staphylococcus pasteuri', 'Corynebacterium imitans', 'Staphylococcus sp. M0911',
    'Moraxella osloensis', 'Streptococcus thermophilus', 'Aeromicrobium erythreum',
    'Roseomonas sp. FDAARGOS_362', 'Agrococcus carbonis', 'Corynebacterium segmentosum',
    'Streptomyces xiamenensis'
]

non_italic = {'sp.', 'sp', 'CS2', 'A12', 'M0911', 'FDAARGOS_362', 'Lophiotremataceae', 'Sordariales', 'Gram', '(+)', '(-)', 'cocci', 'rods'}

def build_perfect_em_table(oos_id, sample_mode, out_docx_path):
    src_template = r"scratch\EM table OOS-262133 and OOS-262134 04SEP2026 (1-Page Master).docx"
    doc = docx.Document(src_template)
    
    # Page setup
    for sec in doc.sections:
        sec.top_margin = Inches(0.20)
        sec.bottom_margin = Inches(0.20)
        sec.left_margin = Inches(0.35)
        sec.right_margin = Inches(0.35)
        
    # Reset title indents to 0 so "Table 1" and "Table 2" are NEVER clipped on the left
    doc.paragraphs[0].text = "Table 1: Read Dates and Incubation Observation"
    doc.paragraphs[2].text = "Table 2: Environmental Monitoring for Analyst and Cleanroom"
    
    for p in doc.paragraphs:
        p.paragraph_format.left_indent = Inches(0)
        p.paragraph_format.first_line_indent = Inches(0)
        p.paragraph_format.right_indent = Inches(0)
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(2)
        for r in p.runs:
            r.font.name = 'Times New Roman'
            r.font.size = Pt(8.0)
            r.bold = True
            
    # Table 1 Setup
    t1 = doc.tables[0]
    t1.rows[0].cells[0].text = "ETX Submission ID"
    t1.rows[0].cells[1].text = "Set-up Analyst and Date"
    t1.rows[0].cells[2].text = "Sampling Site and Location"
    t1.rows[0].cells[3].text = "Plate Reading Analyst (≥ 48H)"
    t1.rows[0].cells[4].text = "CFUs Observed after 48 Hour Incubation at 30–35°C (E001031)"
    t1.rows[0].cells[5].text = "Plate Reading Analyst (NLT 5 days)"
    t1.rows[0].cells[6].text = "CFUs Observed after NLT 5-day Incubation at 20–25°C (E001034)"
    t1.rows[0].cells[7].text = "Microbial Identification"
    
    if sample_mode == "262133":
        t1.rows[1].cells[0].text = "ETX-260914-0493\nETX-260914-0496"
        t1.rows[1].cells[1].text = "SMO\n04SEP 2026"
        t1.rows[1].cells[2].text = "Active air sampling plate"
        t1.rows[1].cells[3].text = "MC"
        t1.rows[1].cells[4].text = "53 CFUs on 143 sec I ISO 8, and 44 CFUs on 143 sec II ISO 8"
        t1.rows[1].cells[5].text = "MC"
        t1.rows[1].cells[6].text = "53 CFUs on 143 sec I ISO 8, and 44 CFUs on 143 sec II ISO 8"
    else:
        t1.rows[1].cells[0].text = "ETX-260914-0498"
        t1.rows[1].cells[1].text = "SMO\n04SEP 2026"
        t1.rows[1].cells[2].text = "Active air sampling plate"
        t1.rows[1].cells[3].text = "MC"
        t1.rows[1].cells[4].text = "17 CFUs on 142 ISO 8"
        t1.rows[1].cells[5].text = "MC"
        t1.rows[1].cells[6].text = "17 CFUs on 142 ISO 8"
        
    def fill_species_cell(cell, sp_list):
        cell.text = ""
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 0.90
        p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
        for idx, sp in enumerate(sp_list):
            if idx > 0:
                p.add_run("\n")
            for w_idx, w in enumerate(sp.split()):
                r = p.add_run(w + (" " if w_idx < len(sp.split())-1 else ""))
                r.font.name = "Times New Roman"
                r.font.size = Pt(7.0)
                if w not in non_italic and not w.startswith("sp"):
                    r.italic = True
                    
    fill_species_cell(t1.rows[1].cells[7], sp_04sep)
    
    # Table 2 Setup
    t2 = doc.tables[1]
    t2.rows[0].cells[0].text = "Environmental Monitoring (EM) Sampling Site"
    t2.rows[0].cells[1].text = "Frequency"
    t2.rows[0].cells[2].text = "Date\n(DDMMMYY)"
    t2.rows[0].cells[3].text = "Analyst"
    t2.rows[0].cells[4].text = "Day /Week(s)"
    t2.rows[0].cells[5].text = "Observation"
    t2.rows[0].cells[6].text = "EM Plate ETX Number"
    t2.rows[0].cells[7].text = "Microbial ID"
    t2.rows[0].cells[8].text = "Notes"
    
    t2.rows[1].cells[0].text = "Weekly Active Air Sampling of Cleanroom 145, 144, 143, 142 (L-Suite) with Processing BSCs"
    t2.rows[5].cells[0].text = "Weekly Surface Sampling of Cleanroom 145, 144, 143, 142 (L-Suite) with Processing BSCs"
    
    # Day/Weeks
    t2.rows[2].cells[4].text = "Week (Before Testing Date)"
    t2.rows[3].cells[4].text = "Week (Of Testing Date)"
    t2.rows[4].cells[4].text = "Week (After Testing Date)"
    t2.rows[6].cells[4].text = "Week (Before Testing Date)"
    t2.rows[7].cells[4].text = "Week (Of Testing Date)"
    t2.rows[8].cells[4].text = "Week (After Testing Date)"
    
    # 04SEP active air
    fill_species_cell(t2.rows[3].cells[7], sp_04sep)
    # 09SEP active air
    fill_species_cell(t2.rows[4].cells[7], sp_09sep)
    
    # Passing surface observation casing & Not Applicable
    for r_idx in [7, 8]:
        t2.rows[r_idx].cells[5].text = "No Growth"
        t2.rows[r_idx].cells[6].text = "Not Applicable"
        t2.rows[r_idx].cells[7].text = "Not Applicable"
        t2.rows[r_idx].cells[8].text = "None"
        
    # Widths - carefully balanced so headers and contents NEVER wrap awkwardly
    t1_widths = [Inches(0.95), Inches(0.70), Inches(0.85), Inches(0.65), Inches(1.20), Inches(0.65), Inches(1.20), Inches(1.60)]
    t2_widths = [Inches(1.05), Inches(0.60), Inches(0.80), Inches(0.50), Inches(0.70), Inches(1.10), Inches(0.95), Inches(1.60), Inches(0.50)]
    
    for t_idx, table in enumerate(doc.tables):
        widths = t1_widths if t_idx == 0 else t2_widths
        table.autofit = False
        table.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
        
        # Generous cell padding so text NEVER touches border lines
        tblCellMar = parse_xml(r'<w:tblCellMar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                               r'<w:top w:w="40" w:type="dxa"/>'
                               r'<w:bottom w:w="40" w:type="dxa"/>'
                               r'<w:left w:w="80" w:type="dxa"/>'
                               r'<w:right w:w="80" w:type="dxa"/>'
                               r'</w:tblCellMar>')
        table._tbl.tblPr.append(tblCellMar)
        
        for r_idx, row in enumerate(table.rows):
            trPr = row._tr.get_or_add_trPr()
            trPr.append(parse_xml(r'<w:cantSplit xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))
            for c_idx, cell in enumerate(row.cells):
                cell.width = widths[c_idx]
                cell.vertical_alignment = docx.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER
                for p in cell.paragraphs:
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.line_spacing = 0.90
                    
                    # Centering: 100% of headers and 100% of data cells are centered
                    if r_idx == 0:
                        p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
                    elif t_idx == 1 and r_idx in [1, 5]:
                        pass # banner left-aligned
                    else:
                        p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
                        
                    for r in p.runs:
                        r.font.name = 'Times New Roman'
                        r.font.size = Pt(7.0)
                        if r_idx == 0 or (t_idx == 1 and r_idx in [1, 5]):
                            r.bold = True
                            
    # Deep XML check: enforce w:sz=14 and w:szCs=14 across all runs in all tables
    for table in doc.tables:
        for r in table._tbl.xpath('.//w:rPr'):
            for sz_node in r.xpath('w:sz'):
                sz_node.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '14')
            for sz_cs in r.xpath('w:szCs'):
                sz_cs.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '14')
                
    doc.save(out_docx_path)
    print(f"Saved docx: {out_docx_path}")

print("build_final_em_tables_v5 loaded.")

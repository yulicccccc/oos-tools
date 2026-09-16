import docx, os

targets = [
    "ScanRDI OOS template.docx",
    "ScanRDI OOS template 0.docx",
    "ScanRDI OOS P1 template.docx",
    "ScanRDI OOS P1 template 0.docx",
]

for tname in targets:
    if not os.path.exists(tname):
        continue
    doc = docx.Document(tname)
    t0 = doc.tables[0]
    
    # Row 6: Section A SOP Info
    # Col 1: SOP / Test Method #
    t0.rows[6].cells[1].text = "SOP / Test Method #:\nMICRO-SOP-12 (16)\nENG-SOP-4 (05)\n"
    
    # Col 2-6: Effective Date
    for c_i in range(2, 7):
        t0.rows[6].cells[c_i].text = "Effective Date: \n24Jul26\n24Jul26"
        
    # Col 7-8: Rev
    for c_i in range(7, 9):
        t0.rows[6].cells[c_i].text = "SOP / Test Method Rev: \nRev: 16\nRev: 05"
        
    # Row 12: Correct SOP followed
    for c_i in range(6, 12):
        t0.rows[12].cells[c_i].text = "Yes, as per MICRO-SOP-12, ENG-SOP-4"
        
    # Row 13: Correct technique followed
    for c_i in range(6, 12):
        t0.rows[13].cells[c_i].text = "Yes, as per MICRO-SOP-12, ENG-SOP-4"
        
    # Check Row 40 in template 0 files
    if "0.docx" in tname:
        cell_text = t0.rows[40].cells[0].text
        new_text = cell_text.replace("SOP 2.600.023, Rapid Scan RDI® Test Using FIFU Method", "MICRO-SOP-12, Rapid Scan RDI® Test using FIFU Method")
        new_text = new_text.replace("SOP 2.600.023, Rapid Scan RDI Test Using FIFU Method", "MICRO-SOP-12, Rapid Scan RDI® Test using FIFU Method")
        new_text = new_text.replace("SOP 2.700.004 (Scan RDI® System – Operations (Standard C3 Quality Check and Microscope Setup and Maintenance)", "ENG-SOP-4 (Scan RDI® System – Operations (Standard C3 Quality Check and Microscope Setup) and Maintenance)")
        new_text = new_text.replace("SOP 2.700.004 (Scan RDI System  Operations (Standard C3 Quality Check and Microscope Setup and Maintenance)", "ENG-SOP-4 (Scan RDI® System – Operations (Standard C3 Quality Check and Microscope Setup) and Maintenance)")
        for c_i in range(len(t0.rows[40].cells)):
            t0.rows[40].cells[c_i].text = new_text

    doc.save(tname)
    print(f"Updated {tname} successfully!")

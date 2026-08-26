import docx

for name in ["EM OOS P1 template.docx", "EM OOS P1 template 0.docx"]:
    doc = docx.Document(name)
    t0 = doc.tables[0]
    
    # Row 6: Section A SOP info
    r6 = t0.rows[6]
    r6.cells[1].text = "SOP / Test Method #:\nMICRO-SOP-2\n"
    r6.cells[2].text = "Effective Date:\n05 Aug 2025\n"
    r6.cells[7].text = "SOP / Test Method Rev:\n15\n"
    
    # Row 12: Correct SOP
    r12 = t0.rows[12]
    for i in range(6, 12):
        r12.cells[i].text = "Yes, as per MICRO-SOP-2"
        
    # Row 13: Correct technique
    r13 = t0.rows[13]
    for i in range(6, 12):
        r13.cells[i].text = "Yes, as per MICRO-SOP-2"
        
    # Row 15: Analyst qualified
    r15 = t0.rows[15]
    for i in range(6, 12):
        r15.cells[i].text = "Yes, the analysts are trained and qualified by quality to perform the test"
        
    # Row 18: Storage
    r18 = t0.rows[18]
    for i in range(6, 12):
        r18.cells[i].text = "Yes, as per MICRO-SOP-2"
        
    doc.save(name)
    print(f"Updated SOP references in {name} to MICRO-SOP-2!")

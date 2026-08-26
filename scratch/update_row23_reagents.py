import docx

for name in ["EM OOS P1 template.docx", "EM OOS P1 template 0.docx"]:
    doc = docx.Document(name)
    t0 = doc.tables[0]
    r23 = t0.rows[23]
    
    # Cells 6 to 10 are Lot Number column
    for i in range(6, 11):
        r23.cells[i].text = "{{ reagent_lot }}"
        
    # Cell 11 is Expiry Date column
    r23.cells[11].text = "{{ reagent_exp }}"
    
    doc.save(name)
    print(f"Updated Row 23 in {name} with {{ reagent_lot }} and {{ reagent_exp }}!")

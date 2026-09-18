import docx

doc_path = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\Tables OOS-261814 GoGoMeds Select (E10747) - ScanRDI.docx"
doc = docx.Document(doc_path)
t = doc.tables[1]

print("Total rows in Table 2:", len(t.rows))

rows_to_check = [0, 7, 9, 11, 12, 13, 15]

for r_idx in rows_to_check:
    row = t.rows[r_idx]
    unique_cells = []
    for c in row.cells:
        txt = c.text.strip().replace("\n", " ")
        if not unique_cells or unique_cells[-1] != txt:
            unique_cells.append(txt)
    print(f"\n--- Row {r_idx} ---")
    print(" | ".join(unique_cells))

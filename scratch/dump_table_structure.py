from docx import Document

doc = Document(r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\EM table OOS-261186 07MAY2026 (2).docx')
with open('scratch/table_structure.txt', 'w', encoding='utf-8') as f:
    for t_idx, table in enumerate(doc.tables):
        f.write(f'=== TABLE {t_idx+1} ===\n')
        for r_idx, row in enumerate(table.rows):
            cells = [c.text.strip().replace('\n', ' ') for c in row.cells]
            f.write(f'Row {r_idx:2d} ({len(cells)} cells):\n')
            for c_idx, c in enumerate(cells):
                f.write(f'  Col {c_idx}: {c}\n')

import os, glob, docx, fitz

bad_patterns = [
    "(IS)", " IS ", "'IS'", '"IS"',
    "(AC)", " AC ", "'AC'", '"AC"',
    "(AO)", " AO ", "'AO'", '"AO"',
    " ES ", "'ES'", '"ES"',
    "Gabriel Anyangwe",
    "Eniola", "ENI",
    "David Karangwa",
    "Jose Ocando",
    "Min Jang",
    "Muralidhar Bythatagari"
]

print("=== 1. AUDITING PYTHON SCRIPTS IN SCRATCH ===")
for pyfile in glob.glob("scratch/*.py"):
    with open(pyfile, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read()
    for bp in bad_patterns:
        if bp in content:
            print(f"  [PYTHON SCRIPT] {pyfile}: contains {repr(bp)}")

print("\n=== 2. AUDITING DESKTOP PDF DELIVERABLES ===")
desktop = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop"
for pdf_file in glob.glob(os.path.join(desktop, "*262098*.pdf")):
    doc = fitz.open(pdf_file)
    for pidx, page in enumerate(doc):
        text = page.get_text()
        for bp in bad_patterns:
            if bp in text:
                print(f"  [PDF] {os.path.basename(pdf_file)} (Page {pidx+1}): contains {repr(bp)}")

print("\n=== 3. AUDITING DESKTOP DOCX DELIVERABLES ===")
for docx_file in glob.glob(os.path.join(desktop, "*262098*.docx")):
    try:
        doc = docx.Document(docx_file)
        full_text = []
        for p in doc.paragraphs:
            full_text.append(p.text)
        for t in doc.tables:
            for r in t.rows:
                for c in r.cells:
                    full_text.append(c.text)
        all_doc_text = " \n ".join(full_text)
        for bp in bad_patterns:
            if bp in all_doc_text:
                print(f"  [DOCX] {os.path.basename(docx_file)}: contains {repr(bp)}")
    except Exception as e:
        print(f"  Error reading {docx_file}: {e}")

print("\n=== AUDIT COMPLETE ===")

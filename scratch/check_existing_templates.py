import docxtpl

templates = [
    "Celsis OOS P1 template.docx",
    "ScanRDI OOS P1 template.docx",
    "USP71 OOS P1 template.docx"
]

for tName in templates:
    try:
        tpl = docxtpl.DocxTemplate(tName)
        vars = sorted(list(tpl.get_undeclared_template_variables()))
        print(f"\n==================== {tName} ({len(vars)} variables) ====================")
        for v in vars:
            print(f"  {{{{ {v} }}}}")
    except Exception as e:
        print(f"Error opening {tName}: {e}")

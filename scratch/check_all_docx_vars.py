import docxtpl

tpl = docxtpl.DocxTemplate("EM OOS P1 template 0.docx")
vars = sorted(list(tpl.get_undeclared_template_variables()))

print(f"Total undeclared variables in EM OOS P1 template 0.docx: {len(vars)}")
for v in vars:
    print(f"  - {v}")

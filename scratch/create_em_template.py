import os
import shutil
import docxtpl

src_docx = "ScanRDI OOS P1 template.docx"
src_docx_0 = "ScanRDI OOS P1 template 0.docx"

dst_docx = "EM OOS P1 template.docx"
dst_docx_0 = "EM OOS P1 template 0.docx"

# 1. Copy
shutil.copyfile(src_docx, dst_docx)
shutil.copyfile(src_docx_0, dst_docx_0)
print(f"Created {dst_docx} and {dst_docx_0}")

# 2. Extract variables
tpl = docxtpl.DocxTemplate(dst_docx)
vars_list = sorted(list(tpl.get_undeclared_template_variables()))

print(f"\nExtracted {len(vars_list)} variables for EM Module:")
for v in vars_list:
    print(f"  - {v}")

with open("scratch/em_variable_contract.txt", "w", encoding="utf-8") as f:
    f.write(f"EM Module Variable Contract ({len(vars_list)} variables):\n")
    for v in vars_list:
        f.write(f"  - {v}\n")

import shutil

src_pdf = "ScanRDI OOS P1 template.pdf"
dst_pdf = "EM OOS P1 template.pdf"

shutil.copyfile(src_pdf, dst_pdf)
print(f"Created {dst_pdf}")

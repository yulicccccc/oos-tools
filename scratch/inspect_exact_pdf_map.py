import pypdf

reader = pypdf.PdfReader("EM OOS P1 template.pdf")
fields = reader.get_fields()

print("Form Fields in EM OOS P1 template.pdf:")
for name, f_obj in fields.items():
    v = f_obj.get('/V', '')
    t = f_obj.get('/T', '')
    print(f"  Field key: '{name}' | Default/Value: {repr(v)}")

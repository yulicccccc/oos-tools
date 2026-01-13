from pypdf import PdfReader

# Load the file
reader = PdfReader("ScanRDI OOS template.pdf")
fields = reader.get_fields()

print("-" * 30)
print(f"FOUND {len(fields)} FIELDS:")
print("-" * 30)

# Print every field name found
if fields:
    for field_name in sorted(fields.keys()):
        print(f"['{field_name}']")
else:
    print("❌ No fields found! The PDF might be flattened.")

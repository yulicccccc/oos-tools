from pypdf import PdfReader

try:
    reader = PdfReader("ScanRDI OOS template.pdf")
    fields = reader.get_fields()
    if fields:
        print("\n✅ SUCCESS! Found these fields:")
        print(list(fields.keys()))
    else:
        print("\n❌ FAIL: The file is a 'Flat PDF'. No text boxes found.")
except Exception as e:
    print(f"\n❌ ERROR: Could not read file. {e}")

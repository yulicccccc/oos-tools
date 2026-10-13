import fitz

doc = fitz.open(r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\CORP-FORM-21 - P1 - 16 Sep 2026.pdf")
all_field_names = set()
for page in doc:
    for w in page.widgets():
        all_field_names.add(w.field_name)

print("Total unique widgets:", len(all_field_names))

# Now let's check our script field names
from scratch.generate_master_oos_262150 import OOS_ID
import scratch.generate_master_oos_262150 as gen

# Let's check which fields from gen.pdf_field_values and cb_whitelists exist
missing_fields = []
# We'll check with the actual fields in generate_master_oos_262150.py

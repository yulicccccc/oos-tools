import os
import docx

desktop_docx = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Desktop\Tables OOS-262237 GoGoMeds Select (E10747) - ScanRDI.docx"
docs_docx = r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\Tables OOS-262237 GoGoMeds Select (E10747) - ScanRDI.docx"

# Load Desktop docx
doc = docx.Document(desktop_docx)

# Check and update rId8 target
fixed = False
for r_id, rel in doc.part.rels.items():
    if 'hyperlink' in rel.reltype:
        print(f"Rel {r_id}: target={rel.target_ref}")
        if '7WzH' in rel.target_ref:
            print(f"Fixing {r_id} from {rel.target_ref} to correct URL!")
            rel._target = "https://etrax.eagleanalytical.com/Submission/Details/QchKBV1D2VrN7y3VezYRkg__"
            fixed = True

if fixed:
    doc.save(docs_docx)
    try:
        doc.save(desktop_docx)
        print("Successfully saved updated docx to Desktop!")
    except Exception as e:
        print(f"Could not save directly to Desktop docx (might be open in Word): {e}")
else:
    print("No 7WzH rel found to fix.")

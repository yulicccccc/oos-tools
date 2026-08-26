import pypdf
import sys
import os

sys.stdout.reconfigure(encoding='utf-8')

# Run test on full PDF map
pdf_map = {
    'Text Field57': '261187',
    'Text Field0': 'Clea S. Garza (written by: Qiyue Chen)',
    'Date Field0': '11May26',
    'Date Field1': '11May26',
    'Date Field2': '11May26',
    'Text Field1': "Environmental Monitoring",
    'Text Field2': 'ETX-260518-0273',
    'Text Field3': "Clea S. Garza (CGS)\n(Surface Sampling Plate Setup)\n\nSimin Mohammad\n(Surface Sampling Plate Reader)",
    'Text Field4': "ScanC/O CGS E001309 S1 11MAY2026",
    'Text Field5': "Plate",
    'Text Field6': "ScanC/O CGS E001309 S1 11MAY2026",
    'Text Field7': "The CFU count for the environmental monitoring plate exceeded the action level.",
    'Text Field8': "2.600.002",
    'Text Field11': "Action Level: >= 1 CFU/Plate",
    'Text Field13': "Yes, Clea S. Garza and Simin Mohammad were interviewed comprehensively.",
    'Text Field14': "N/A",
    'Text Field15': "Yes, as per SOP 2.600.002",
    'Text Field16': "Yes, as per SOP 2.600.002",
    'Text Field17': "Yes, Information is available in EagleTrax under ETX-260518-0273.",
    'Text Field18': "Yes, the analysts are trained and qualified by quality to perform the test.",
    'Text Field19': "Not Applicable",
    'Text Field20': "Not Applicable",
    'Text Field21': "Yes, as per SOP 2.600.002",
    'Text Field24': "Not Applicable",
    'Text Field25': "Not Applicable",
    'Text Field26': "Not Applicable",
    'Text Field27': "Not Applicable",
    'Text Field28': "Not Applicable",
    'Text Field29': "Not Applicable",
    'Text Field30': "Please see below",
    'Text Field31': "Please see below",
    'Text Field34': "Not Applicable",
    'Text Field35': "Not Applicable",
    'Text Field36': "Not Applicable",
    'Text Field37': "Not Applicable",
    'Text Field38': "Not Applicable",
    'Text Field39': "Not Applicable",
    'Text Field40': "Not Applicable",
    'Text Field41': "Not Applicable",
    'Text Field42': "Not Applicable",
    'Text Field45': "Not Applicable",
    'Text Field46': "Not Applicable",
    'Text Field47': "Not Applicable",
    'Text Field53': "Maryam Naeem",
    'Text Field54': "Kathan Parikh"
}

# Add standard Yes checkboxes
yes_boxes = [4, 9, 10, 13, 16, 19, 24, 27, 28, 32, 34, 38, 42, 43, 48, 51, 52, 55, 60, 63, 66, 69, 72, 73, 78, 87, 79]
for b_num in yes_boxes:
    pdf_map[f'Check Box{b_num}'] = '/Yes'

writer = pypdf.PdfWriter(clone_from="EM OOS P1 template.pdf")
for page in writer.pages:
    writer.update_page_form_field_values(page, pdf_map)

out_pdf = "scratch/test_audit_complete.pdf"
with open(out_pdf, "wb") as f:
    writer.write(f)

print(f"Generated complete PDF test file: {out_pdf}")

# Verify generated PDF fields
reader = pypdf.PdfReader(out_pdf)
fields = reader.get_fields()
print(f"Total fields populated: {len(fields)}")
print("Key checklist items in output PDF:")
for check_key in ['Text Field0', 'Text Field2', 'Text Field3', 'Text Field8', 'Text Field13', 'Text Field17', 'Text Field18', 'Text Field53', 'Text Field54']:
    print(f"  {check_key}: {repr(fields[check_key].get('/V'))}")

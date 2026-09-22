import fitz
import os

base_pdf = r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS\scratch\updated_9p_ncr.pdf'
out_pdf_docs = r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\N26487 - Scan RDI Instrument E001230 Software Crash.pdf'
out_pdf_oos = r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS\N26487 - Scan RDI Instrument E001230 Software Crash.pdf'

out_pdf_docs_final = r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\N26487 - Scan RDI Instrument E001230 Software Crash - Final.pdf'
out_pdf_oos_final = r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS\N26487 - Scan RDI Instrument E001230 Software Crash - Final.pdf'

doc_src = fitz.open(base_pdf)
new_doc = fitz.open()
new_doc.insert_pdf(doc_src, from_page=0, to_page=1)

p1 = new_doc[0]

# Areas to white out on Page 1 (inside cells, leaving borders intact)
whiteout_rects_p1 = [
    fitz.Rect(484, 202, 524, 212),  # NCR#
    fitz.Rect(115, 225, 227, 235),  # Initiator
    fitz.Rect(254, 225, 523, 235),  # Date
    fitz.Rect(131, 236, 227, 247),  # Dept
    fitz.Rect(147, 248, 227, 258),  # Equip ID
    fitz.Rect(230, 245, 523, 258),  # Sample Name
    fitz.Rect(133, 259, 227, 270),  # Sample ID
    fitz.Rect(322, 259, 523, 270),  # Lot #
    fitz.Rect(78, 292, 523, 368),   # Nonconformance Desc
    fitz.Rect(178, 381, 523, 406),  # Cause
    fitz.Rect(115, 408, 210, 423),  # Other text (N/A 21Aug26 SMO)
    fitz.Rect(78, 443, 523, 512),   # Corrective Action
    fitz.Rect(176, 576, 523, 634),  # Comments
    fitz.Rect(78, 656, 298, 682),   # Prepared by
    fitz.Rect(78, 693, 298, 719),   # Reviewed by
]

for r in whiteout_rects_p1:
    p1.draw_rect(r, color=None, fill=(1, 1, 1))

# Process Page 1 widgets
for w in list(p1.widgets()):
    fn = w.field_name
    if fn == 'Text Field20':
        p1.delete_widget(w)
        continue
    elif fn == 'Text Field10':
        w.field_value = ''
    elif fn == 'Text Field0':
        w.text_fontsize = 8.0
        w.field_value = '26487'
    elif fn == 'Text Field1':
        w.text_fontsize = 8.0
        w.field_value = 'Qiyue Chen'
    elif fn == 'Date Field0':
        w.text_fontsize = 8.0
        w.field_value = '18-Sep-2026'
    elif fn == 'Text Field2':
        w.text_fontsize = 8.0
        w.field_value = 'Microbiology'
    elif fn == 'Text Field3':
        # Equipment ID #
        w.text_fontsize = 8.0
        w.field_value = 'E001230'
    elif fn == 'Text Field4':
        # Sample ID # - 3 specific samples
        w.rect = fitz.Rect(130.0, 258.0, 227.5, 271.5)
        w.text_fontsize = 4.0
        w.field_value = 'ETX-260915-0653, ETX-260916-0044,\nETX-260916-0374'
    elif fn == 'Text Field5':
        # Sample Name / Material Description - 3 specific samples
        w.text_fontsize = 5.2
        w.field_value = 'Semaglutide/Cyanocobalamin; Tirzepatide 22mg/mL-Pyridoxine HCl; SEMAGLUTIDE/B12'
    elif fn == 'Text Field6':
        # Sample/Material Lot # - 3 specific lot numbers
        w.text_fontsize = 5.8
        w.field_value = '0216202602, 22666, LG342010613'
    elif fn == 'Text Field7':
        # Nonconformance Description with complete sample/client/lot details
        w.text_fontsize = 5.2
        w.field_value = (
            "On 17-Sep-2026, during routine Scan RDI sterility testing operations, an unexpected software crash occurred on the Scan RDI instrument (Equipment ID: E001230) while running session '17 Sep 2026 - 3'. Following the crash, third-shift analysts and Microbiology Laboratory Supervisor Robin Seymour attempted to recover the session; however, the instrument displayed system error messages stating 'Unable to Create New Session Unknown Error' and 'Unable to Load Session Unknown Error', preventing completion of the verification and reading process.\n\n"
            "On 18-Sep-2026, Eagle IT Coordinator Kevin Torres conducted a technical assessment of instrument E001230 and confirmed that the session file was corrupted, exporting as only 7 KB, and the corresponding raw acquisition dataset within the D:\\ directory was completely empty and unrecoverable.\n\n"
            "A total of three (3) sterility test sample submissions were actively processed in this corrupted session and impacted by data loss:\n"
            "• ETX-260915-0653 | Client: ANG Labs (E73555) | Sample: Semaglutide/Cyanocobalamin | Lot#: 0216202602\n"
            "• ETX-260916-0044 | Client: Southend Pharmacy (E19207) | Sample: Tirzepatide 22mg/mL-Pyridoxine HCl 4mg/mL (1 mL) | Lot#: 22666\n"
            "• ETX-260916-0374 | Client: Optimal Balance Pharmacy (E19193) | Sample: SEMAGLUTIDE/ B12 1.2 MG/ 500 MCG/ML (2 ML) Inj | Lot#: LG342010613\n\n"
            "No valid test results could be generated or verified for these submissions from the corrupted run. Work Order WO-260470 was promptly issued to Engineering/IT to formally investigate the software failure and initiate technical remediation."
        )
    elif fn == 'Text Field8':
        # Description of cause(s)
        w.text_fontsize = 6.8
        w.field_value = 'Scan RDI software crash and session file corruption on instrument E001230, resulting in unrecoverable raw acquisition data in the D:\\ directory.'
    elif fn == 'Check Box0':
        # Equipment/System
        w.field_value = 'Yes'
    elif fn in ['Check Box1', 'Check Box2', 'Check Box3', 'Check Box4']:
        w.field_value = ''
    elif fn == 'Text Field9':
        # Description of corrective action(s)
        w.text_fontsize = 6.0
        w.field_value = (
            "1. Work Order WO-260470 was issued to Engineering and IT to investigate and resolve the software crash, clear corrupted temporary files, and evaluate database integrity on instrument E001230.\n"
            "2. Instrument E001230 was taken out of service for live sample testing pending completion and sign-off of WO-260470.\n"
            "3. Client Care was instructed to request additional sample vials for retesting from ANG Labs (ETX-260915-0653) and Southend Pharmacy (ETX-260916-0044). Additional sample for ETX-260916-0374 (Optimal Balance Pharmacy) was received and is in-process with the third shift for retesting on an alternate qualified Scan RDI unit."
        )
    elif fn == 'Check Box5':
        w.field_value = 'Yes'
    elif fn in ['Check Box6', 'Check Box7', 'Check Box8', 'Check Box9']:
        w.field_value = ''
    elif fn == 'Text Field11':
        # Risk Assessment / Comments
        w.text_fontsize = 6.5
        w.field_value = (
            "Impact is isolated to the 3 sample submissions processed in the corrupted session (ETX-260915-0653, ETX-260916-0044, ETX-260916-0374). No impact to other testing sessions or alternate operational Scan RDI instruments (E002017, E002225). Additional vials requested / retesting initiated. Instrument E001230 remains under observation and repair under WO-260470."
        )
    elif fn == 'Text Field12':
        w.text_fontsize = 8.5
        w.field_value = 'Qiyue Chen'
    elif fn == 'Text Field13':
        w.text_fontsize = 8.5
        w.field_value = 'Robin Seymour'
    w.update()

# Page 2
p2 = new_doc[1]
p2.draw_rect(fitz.Rect(484, 202, 524, 212), color=None, fill=(1, 1, 1))

for w in p2.widgets():
    fn = w.field_name
    if fn == 'Text Field0':
        w.text_fontsize = 8.0
        w.field_value = '26487'
    elif fn == 'Text Field14':
        w.text_fontsize = 8.0
        w.field_value = '2'
    elif fn == 'Text Field15':
        w.text_fontsize = 8.0
        w.field_value = '2'
    elif fn == 'Text Field16':
        w.text_fontsize = 8.0
        w.field_value = '2'
    elif fn == 'Text Field17':
        w.text_fontsize = 8.0
        w.field_value = '8'
    elif fn == 'Check Box10':
        w.field_value = ''
    elif fn == 'Check Box11':
        w.field_value = 'Yes'
    elif fn == 'Check Box12':
        w.field_value = ''
    elif fn == 'Text Field18':
        w.text_fontsize = 8.5
        w.field_value = 'Robin Seymour'
    elif fn == 'Text Field19':
        w.field_value = ''
    w.update()

# Save with fallback if file is locked
saved_paths = []
try:
    new_doc.save(out_pdf_docs)
    saved_paths.append(out_pdf_docs)
except Exception as e:
    print('Could not save to', out_pdf_docs, '(locked). Saving to Final copy.')

try:
    new_doc.save(out_pdf_oos)
    saved_paths.append(out_pdf_oos)
except Exception as e:
    print('Could not save to', out_pdf_oos, '(locked). Saving to Final copy.')

new_doc.save(out_pdf_docs_final)
saved_paths.append(out_pdf_docs_final)
new_doc.save(out_pdf_oos_final)
saved_paths.append(out_pdf_oos_final)

new_doc.close()
doc_src.close()

# Re-render PNGs
d = fitz.open(out_pdf_docs_final)
pix1 = d[0].get_pixmap(dpi=150)
pix1.save(r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS\scratch\clean_n26487_p1.png')

pix2 = d[1].get_pixmap(dpi=150)
pix2.save(r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS\scratch\clean_n26487_p2.png')
print('Successfully regenerated clean N26487 PDF & PNGs with full sample info.')
print('Saved to:', saved_paths)

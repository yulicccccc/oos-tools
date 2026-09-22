import fitz
import os

base_pdf = r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS\scratch\updated_9p_ncr.pdf'
out_pdf_docs = r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\N26487 - Scan RDI Instrument E001230 Software Crash.pdf'
out_pdf_oos = r'C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS\N26487 - Scan RDI Instrument E001230 Software Crash.pdf'

doc_src = fitz.open(base_pdf)
new_doc = fitz.open()
new_doc.insert_pdf(doc_src, from_page=0, to_page=1)

# Page 1
p1 = new_doc[0]
for w in p1.widgets():
    fn = w.field_name
    if fn == 'Text Field0':
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
        # Sample ID #
        w.text_fontsize = 6.2
        w.field_value = 'ETX-260915-0653, ETX-260916-0044, ETX-260916-0374'
    elif fn == 'Text Field5':
        # Sample Name / Material Description
        w.text_fontsize = 7.5
        w.field_value = 'Scan RDI Sterility Testing Samples (Session 17 Sep 2026 - 3)'
    elif fn == 'Text Field6':
        # Sample/Material Lot #
        w.text_fontsize = 7.5
        w.field_value = 'N/A (Multiple Client Lots)'
    elif fn == 'Text Field7':
        # Nonconformance Description
        w.text_fontsize = 5.7
        w.field_value = (
            "On 17-Sep-2026, during routine Scan RDI sterility testing operations, an unexpected software crash occurred on the Scan RDI instrument (Equipment ID: E001230) while running session '17 Sep 2026 - 3'. Following the crash, third-shift analysts and Microbiology Laboratory Supervisor Robin Seymour attempted to recover the session; however, the instrument displayed system error messages stating 'Unable to Create New Session Unknown Error' and 'Unable to Load Session Unknown Error', preventing completion of the verification and reading process.\n\n"
            "On 18-Sep-2026, Eagle IT Coordinator Kevin Torres conducted a technical assessment of instrument E001230 and confirmed that the session file was corrupted, exporting as only 7 KB, and the corresponding raw acquisition dataset within the D:\\ directory was completely empty and unrecoverable.\n\n"
            "A total of three (3) sterility test sample submissions were actively processed in this corrupted session and impacted by data loss: ETX-260915-0653, ETX-260916-0044, and ETX-260916-0374. No valid test results could be generated or verified for these submissions from the corrupted run. Work Order WO-260470 was promptly issued to Engineering/IT to formally investigate the software failure and initiate technical remediation."
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
    elif fn == 'Text Field20':
        w.field_value = ''
    elif fn == 'Text Field9':
        # Description of corrective action(s)
        w.text_fontsize = 6.2
        w.field_value = (
            "1. Work Order WO-260470 was issued to Engineering and IT to investigate and resolve the software crash, clear corrupted temporary files, and evaluate database integrity on instrument E001230.\n"
            "2. Instrument E001230 was taken out of service for live sample testing pending completion and sign-off of WO-260470.\n"
            "3. Sample re-testing was coordinated for the three affected submissions: ETX-260916-0374 was received and in-process with the third shift for re-testing on an alternate qualified Scan RDI unit; client notifications and rerun procedures were initiated for ETX-260915-0653 and ETX-260916-0044 per SOP."
        )
    elif fn == 'Check Box5':
        w.field_value = 'Yes'
    elif fn in ['Check Box6', 'Check Box7', 'Check Box8', 'Check Box9']:
        w.field_value = ''
    elif fn == 'Text Field10':
        w.field_value = ''
    elif fn == 'Text Field11':
        # Risk Assessment / Comments
        w.text_fontsize = 6.5
        w.field_value = (
            "Impact is isolated to the 3 sample submissions processed in the corrupted session (ETX-260915-0653, ETX-260916-0044, ETX-260916-0374). No impact to other testing sessions or alternate operational Scan RDI instruments (E002017, E002225). Rerun procedures were immediately initiated. Instrument E001230 remains under observation and repair under WO-260470."
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

new_doc.save(out_pdf_docs)
new_doc.save(out_pdf_oos)
new_doc.close()
doc_src.close()
print('Successfully generated N26487 PDF to:')
print(' -', out_pdf_docs)
print(' -', out_pdf_oos)

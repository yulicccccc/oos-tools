import os
import win32clipboard

html_content = '''<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  body { font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #000; line-height: 1.4; }
  table { border-collapse: collapse; width: 100%; font-family: Calibri, Arial, sans-serif; font-size: 10pt; margin: 10px 0 15px 0; }
  th { background-color: #FFFF00 !important; color: #000; font-weight: bold; text-align: center; border: 1px solid #B0B0B0; padding: 6px 8px; }
  td { border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top; }
</style>
</head>
<body style="font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #000; line-height: 1.4;">
<!--StartFragment-->
<p style="margin: 0 0 8px 0; font-family: Calibri, sans-serif; font-size: 11pt;">Good morning @Simin Mohammad,</p>
<p style="margin: 0 0 12px 0; font-family: Calibri, sans-serif; font-size: 11pt;">I have reviewed this OOS. Please find the summary of edits in the table below. All required updates, including the amended EM table and all form field revisions, have already been completed on your behalf and assembled into the attached finalized 7-page PDF package.</p>
<table border="1" bordercolor="#B0B0B0" cellpadding="6" cellspacing="0" style="border-collapse: collapse; width: 100%; font-family: Calibri, sans-serif; font-size: 10pt; margin: 10px 0 15px 0; text-align: left;">
  <thead>
    <tr>
      <th bgcolor="#FFFF00" style="background-color: #FFFF00 !important; width: 20%; color: #000; font-weight: bold; text-align: center; border: 1px solid #B0B0B0; padding: 6px 8px;">Section</th>
      <th bgcolor="#FFFF00" style="background-color: #FFFF00 !important; width: 30%; color: #000; font-weight: bold; text-align: center; border: 1px solid #B0B0B0; padding: 6px 8px;">Unreviewed</th>
      <th bgcolor="#FFFF00" style="background-color: #FFFF00 !important; width: 32%; color: #000; font-weight: bold; text-align: center; border: 1px solid #B0B0B0; padding: 6px 8px;">Reviewed</th>
      <th bgcolor="#FFFF00" style="background-color: #FFFF00 !important; width: 18%; color: #000; font-weight: bold; text-align: center; border: 1px solid #B0B0B0; padding: 6px 8px;">Impact</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; font-weight: bold; vertical-align: top;">EM Table Attachment &amp; Formatting</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">The EM Table was omitted from the PDF submission. In the draft Word document, Table 1 mistakenly listed ETX-260526-0461 and Talaromyces purpurogenus instead of the actual isolates, Table 2 contained formatting inconsistencies such as unclosed parentheses in timing descriptions, and multi-organism entries ran together without line breaks.</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">Amended both tables to reflect the correct ETX submission ID ETX-260518-0250, updated the microbial identification to Penicillium decumbens, Cladosporium tenuissimum, Cladosporium langeronii, and Cladosporium halotolerans, corrected all timing descriptions, properly formatted multi-organism line breaks, and attached the amended 1-page table as Page 7 to complete the 7-page package.</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">Document Assembly &amp; Technical Uniformity Correction.<br><i>Completed on analyst's behalf.</i></td>
    </tr>
    <tr>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; font-weight: bold; vertical-align: top;">Initiator Name</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">Listed as &quot;Simin Mohammad (written by Simin Mohammad)&quot;.</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">Revised to &quot;Simin Mohammad&quot; to remove redundant parenthetical phrasing.</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">Clerical Correction.</td>
    </tr>
    <tr>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; font-weight: bold; vertical-align: top;">Name of Analyst who Performed the Test</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">Listed setup and reader analysts running together with lowercase &quot;weekly&quot; and plural &quot;Readers&quot;: &quot;Simin Mohammad (weekly Active air Sampling Plate Setup) Maraya Chukwumerije &amp; Sophia Santamaria(Weekly active air Sampling Plate Readers)&quot;.</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">Formatted into distinct, properly capitalized entries separating Simin Mohammad as the setup analyst from Maraya Chukwumerije and Sophia Santamaria as individual plate readers:<br>Simin Mohammad (Weekly Active Air Sampling Plate Setup)<br>Maraya Chukwumerije (Weekly Active Air Sampling Plate Reader)<br>Sophia Santamaria (Weekly Active Air Sampling Plate Reader)</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">Formatting &amp; Administrative Correction.</td>
    </tr>
    <tr>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; font-weight: bold; vertical-align: top;">SOP Reference &amp; Effective Date</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">Listed as &quot;20600.002&quot; Revision 15 with Effective Date &quot;05AUG2025&quot;.</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">Revised to &quot;MICRO-SOP-2&quot; Revision 16 with Effective Date &quot;23-Jul-2026&quot;.</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">Compliance &amp; Document Control Correction.</td>
    </tr>
    <tr>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; font-weight: bold; vertical-align: top;">Limits / Specification</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">Listed as &quot;Action level &gt;10&quot;.</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">Revised to &quot;Action level: &gt;= 10 CFU/Plate&quot;.</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">Technical Specification Correction.</td>
    </tr>
    <tr>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; font-weight: bold; vertical-align: top;">Section B SOP Citations</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">SOP reference in comments contained a typographical error listed as &quot;SOP 2.600.00&quot; missing the trailing digit.</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">Revised all instances to &quot;MICRO-SOP-2&quot;.</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">SOP Citation Correction.</td>
    </tr>
    <tr>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; font-weight: bold; vertical-align: top;">Analyst Interview Comment</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">&quot;Yes, analysts Simin Mohammad, Maraya Chukwumerije, &amp; Sophia Santamaria were comprehensively interviewed.&quot;</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">Revised ampersand to &quot;and&quot;: &quot;Yes, analysts Simin Mohammad, Maraya Chukwumerije, and Sophia Santamaria were comprehensively interviewed.&quot;</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">Grammatical Correction.</td>
    </tr>
    <tr>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; font-weight: bold; vertical-align: top;">Equipment &amp; Calibration Details</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">Equipment IDs and calibration dates were running together without proper spacing: &quot;Incubator E001034(Sensor E001501)Incubator E001031 (Sensor E001505)&quot; and &quot;Aug 2026 Feb 2027  Aug 2026 Feb 2027&quot;.</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">Revised with clean line breaks to clearly delineate Incubator E001034 with Sensor E001501 and Incubator E001031 with Sensor E001505 alongside their respective calibration dates of August 2026 and February 2027.</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">Equipment Traceability &amp; Readability Correction.</td>
    </tr>
    <tr>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; font-weight: bold; vertical-align: top;">Phase I Summary – Setup &amp; Incubation Narrative</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">Contained a spelling error in reader Sophia Santamaria's surname (&quot;Sophia Sanatamaria&quot;), grammatical disagreement regarding incubator functionality (&quot;incubators were verified&quot;), and an inaccurate reference to biosafety cabinet locations instead of cleanroom suites.</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">Corrected the analyst's surname to Sophia Santamaria, corrected grammatical phrasing to indicate incubator functionality was verified, updated cleanroom locations to Suite 115 (ISO 8), Suite 115A (ISO 7), and Suite 115B (ISO 7), and cited MICRO-SOP-2 and MICRO-SOP-9.</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">Grammatical &amp; Scope Correction.</td>
    </tr>
    <tr>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; font-weight: bold; vertical-align: top;">Phase I Summary – Monthly Cleaning &amp; Disinfection Narrative</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">Contained contradictory monthly cleaning dates of 22 Feb 2026 and 26APR2026 within the same sentence, alongside a missing period and legacy SOP citations.</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">Reconciled the cleaning date to 26-Apr-2026 performed by Tamiru Kotisso and Cuong Du, inserted the missing period before the H2O2 indicators statement, and updated the procedure reference to MICRO-SOP-9.</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">Chronological Accuracy &amp; Compliance Correction.</td>
    </tr>
    <tr>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; font-weight: bold; vertical-align: top;">Phase I Summary – Root Cause Statement</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">Stated that the OOS result observed for the Environmental Monitoring (EM) Settling Sampling plate may be attributed to a potential analyst error.</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">Corrected the sampling plate type from Settling Sampling plate to Active Air Sampling plate to accurately reflect the test performed, while maintaining the determination that the result may be attributed to a potential analyst error, noting that the contamination was transient in nature and no corrective actions are necessary.</td>
      <td style="border: 1px solid #B0B0B0; padding: 6px 8px; vertical-align: top;">Sampling Type &amp; Technical Accuracy Correction.</td>
    </tr>
  </tbody>
</table>
<p style="margin: 12px 0 6px 0; font-family: Calibri, sans-serif; font-size: 11pt;">@Simin Mohammad please review and sign the attached finalized 7-page PDF package.</p>
<p style="margin: 0; font-family: Calibri, sans-serif; font-size: 11pt;">Thanks,</p>
<!--EndFragment-->
</body>
</html>'''

header_tmpl = 'Version:0.9\r\nStartHTML:{:08d}\r\nEndHTML:{:08d}\r\nStartFragment:{:08d}\r\nEndFragment:{:08d}\r\n'
dummy = header_tmpl.format(0, 0, 0, 0)
header_len = len(dummy.encode('utf-8'))

start_html = header_len
end_html = header_len + len(html_content.encode('utf-8'))

start_frag_idx = html_content.find('<!--StartFragment-->') + len('<!--StartFragment-->')
start_frag = header_len + len(html_content[:start_frag_idx].encode('utf-8'))

end_frag_idx = html_content.find('<!--EndFragment-->')
end_frag = header_len + len(html_content[:end_frag_idx].encode('utf-8'))

final_header = header_tmpl.format(start_html, end_html, start_frag, end_frag)
payload = final_header.encode('utf-8') + html_content.encode('utf-8')

plain_text = '''Good morning @Simin Mohammad,

I have reviewed this OOS. Please find the summary of edits in the table below. All required updates, including the amended EM table and all form field revisions, have already been completed on your behalf and assembled into the attached finalized 7-page PDF package.

Section | Unreviewed | Reviewed | Impact
---------------------------------------------------------------------------------------------------------
EM Table Attachment & Formatting:
  Unreviewed: The EM Table was omitted from the PDF submission. In the draft Word document, Table 1 mistakenly listed ETX-260526-0461 and Talaromyces purpurogenus instead of the actual isolates, Table 2 contained formatting inconsistencies such as unclosed parentheses in timing descriptions, and multi-organism entries ran together without line breaks.
  Reviewed: Amended both tables to reflect the correct ETX submission ID ETX-260518-0250, updated the microbial identification to Penicillium decumbens, Cladosporium tenuissimum, Cladosporium langeronii, and Cladosporium halotolerans, corrected all timing descriptions, properly formatted multi-organism line breaks, and attached the amended 1-page table as Page 7 to complete the 7-page package.
  Impact: Document Assembly & Technical Uniformity Correction. Completed on analyst's behalf.

Initiator Name:
  Unreviewed: Listed as "Simin Mohammad (written by Simin Mohammad)".
  Reviewed: Revised to "Simin Mohammad" to remove redundant parenthetical phrasing.
  Impact: Clerical Correction.

Name of Analyst who Performed the Test:
  Unreviewed: Listed setup and reader analysts running together with lowercase "weekly" and plural "Readers": "Simin Mohammad (weekly Active air Sampling Plate Setup) Maraya Chukwumerije & Sophia Santamaria(Weekly active air Sampling Plate Readers)".
  Reviewed: Formatted into distinct, properly capitalized entries separating Simin Mohammad as the setup analyst from Maraya Chukwumerije and Sophia Santamaria as individual plate readers:
  Simin Mohammad (Weekly Active Air Sampling Plate Setup)
  Maraya Chukwumerije (Weekly Active Air Sampling Plate Reader)
  Sophia Santamaria (Weekly Active Air Sampling Plate Reader)
  Impact: Formatting & Administrative Correction.

SOP Reference & Effective Date:
  Unreviewed: Listed as "20600.002" Revision 15 with Effective Date "05AUG2025".
  Reviewed: Revised to "MICRO-SOP-2" Revision 16 with Effective Date "23-Jul-2026".
  Impact: Compliance & Document Control Correction.

Limits / Specification:
  Unreviewed: Listed as "Action level >10".
  Reviewed: Revised to "Action level: >= 10 CFU/Plate".
  Impact: Technical Specification Correction.

Section B SOP Citations:
  Unreviewed: SOP reference in comments contained a typographical error listed as "SOP 2.600.00" missing the trailing digit.
  Reviewed: Revised all instances to "MICRO-SOP-2".
  Impact: SOP Citation Correction.

Analyst Interview Comment:
  Unreviewed: "Yes, analysts Simin Mohammad, Maraya Chukwumerije, & Sophia Santamaria were comprehensively interviewed."
  Reviewed: Revised ampersand to "and": "Yes, analysts Simin Mohammad, Maraya Chukwumerije, and Sophia Santamaria were comprehensively interviewed."
  Impact: Grammatical Correction.

Equipment & Calibration Details:
  Unreviewed: Equipment IDs and calibration dates were running together without proper spacing: "Incubator E001034(Sensor E001501)Incubator E001031 (Sensor E001505)" and "Aug 2026 Feb 2027  Aug 2026 Feb 2027".
  Reviewed: Revised with clean line breaks to clearly delineate Incubator E001034 with Sensor E001501 and Incubator E001031 with Sensor E001505 alongside their respective calibration dates of August 2026 and February 2027.
  Impact: Equipment Traceability & Readability Correction.

Phase I Summary – Setup & Incubation Narrative:
  Unreviewed: Contained a spelling error in reader Sophia Santamaria's surname ("Sophia Sanatamaria"), grammatical disagreement regarding incubator functionality ("incubators were verified"), and an inaccurate reference to biosafety cabinet locations instead of cleanroom suites.
  Reviewed: Corrected the analyst's surname to Sophia Santamaria, corrected grammatical phrasing to indicate incubator functionality was verified, updated cleanroom locations to Suite 115 (ISO 8), Suite 115A (ISO 7), and Suite 115B (ISO 7), and cited MICRO-SOP-2 and MICRO-SOP-9.
  Impact: Grammatical & Scope Correction.

Phase I Summary – Monthly Cleaning & Disinfection Narrative:
  Unreviewed: Contained contradictory monthly cleaning dates of 22 Feb 2026 and 26APR2026 within the same sentence, alongside a missing period and legacy SOP citations.
  Reviewed: Reconciled the cleaning date to 26-Apr-2026 performed by Tamiru Kotisso and Cuong Du, inserted the missing period before the H2O2 indicators statement, and updated the procedure reference to MICRO-SOP-9.
  Impact: Chronological Accuracy & Compliance Correction.

Phase I Summary – Root Cause Statement:
  Unreviewed: Stated that the OOS result observed for the Environmental Monitoring (EM) Settling Sampling plate may be attributed to a potential analyst error.
  Reviewed: Corrected the sampling plate type from Settling Sampling plate to Active Air Sampling plate to accurately reflect the test performed, while maintaining the determination that the result may be attributed to a potential analyst error, noting that the contamination was transient in nature and no corrective actions are necessary.
  Impact: Sampling Type & Technical Accuracy Correction.

@Simin Mohammad please review and sign the attached finalized 7-page PDF package.

Thanks,
'''

win32clipboard.OpenClipboard()
try:
    win32clipboard.EmptyClipboard()
    CF_HTML = win32clipboard.RegisterClipboardFormat("HTML Format")
    win32clipboard.SetClipboardData(CF_HTML, payload)
    win32clipboard.SetClipboardText(plain_text)
    print("SUCCESS: Copied rich HTML table and complete plain text directly into Windows Clipboard!")
finally:
    win32clipboard.CloseClipboard()

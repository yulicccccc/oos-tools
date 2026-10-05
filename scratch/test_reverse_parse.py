import re

def parse_completed_email(text, email_type="incompatible"):
    # Strip HTML tags if HTML is passed
    clean = re.sub(r'(?i)<br\s*/?>', '\n', text)
    clean = re.sub(r'(?i)</(?:p|div|tr|li|h[1-6])>', '\n', clean)
    clean = re.sub(r'<[^>]+>', '', clean)
    clean = clean.replace('&nbsp;', ' ').replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>').replace('&quot;', '"')
    
    lines = [l.strip() for l in clean.split('\n') if l.strip()]
    
    # Identify sample submission ETXs:
    # A true submission ETX is:
    # 1) The entire line is just an ETX (or starts with ETX), NOT preceded by "Method suitability", "under", "prior", "retest", etc.
    # 2) Or followed closely by "Sample:" / "Sample name:" / "Lot:"
    etx_indices = []
    for idx, line in enumerate(lines):
        # Must not be inside notes/history
        if re.search(r'\b(?:method suitability|prior|under|retest|lot)\b', line, re.I):
            continue
        m = re.search(r'\b(ETX-\d{6}-\d{4})\b', line, re.I)
        if m:
            # check if line is essentially just the ETX, or line has no colon
            # or next line has "Sample:"
            is_submission = False
            if re.match(r'^(?:Submission\s*ID(?:\s*\(ETX\))?\s*:\s*)?ETX-\d{6}-\d{4}$', line, re.I):
                is_submission = True
            elif idx + 1 < len(lines) and re.match(r'^(?:Sample|Lot|Dosage|Test\s*date)', lines[idx + 1], re.I):
                is_submission = True
            elif idx + 2 < len(lines) and re.match(r'^(?:Sample|Lot|Dosage|Test\s*date)', lines[idx + 2], re.I):
                is_submission = True
            elif line.startswith(m.group(1)):
                is_submission = True
                
            if is_submission:
                etx_indices.append((idx, m.group(1).upper()))
            
    if not etx_indices:
        return None
        
    samples = []
    for i, (etx_idx, etx_id) in enumerate(etx_indices):
        client_name = ""
        if etx_idx > 0:
            prev_line = lines[etx_idx - 1]
            ignore_keywords = ["after review", "hi @", "client care", "attention", "attn:", "subject:", "to:", "cc:", "scan rdi testing:"]
            if not any(k in prev_line.lower() for k in ignore_keywords) and not prev_line.lower().startswith("conclusion:") and not prev_line.lower().startswith("---"):
                client_name = prev_line
                
        start_line = etx_idx
        end_line = etx_indices[i + 1][0] if i + 1 < len(etx_indices) else len(lines)
        sample_lines = lines[start_line:end_line]
        sample_text = "\n".join(sample_lines)
        
        sample = {
            "clientName": client_name,
            "submissionId": etx_id,
            "sampleName": "",
            "lotNumber": "",
            "dosageForm": "",
            "testDate": "",
            "analystNotes": "",
            "processingNotes": "",
            "readingNotes": "",
            "sampleHistory": "",
            "conclusion": "",
            "customConclusion": "",
            "isFromCompletedEmail": True
        }
        
        # Extract fields
        m_sample = re.search(r'(?:Sample(?:\s*name)?)\s*:\s*([^\n\r]+)', sample_text, re.I)
        if m_sample: sample["sampleName"] = m_sample.group(1).strip()
        
        m_lot = re.search(r'(?:Lot(?:\s*#)?)\s*:\s*([^\n\r]+)', sample_text, re.I)
        if m_lot: sample["lotNumber"] = m_lot.group(1).strip()
        
        m_dosage = re.search(r'Dosage\s*form\s*:\s*([^\n\r]+)', sample_text, re.I)
        if m_dosage: sample["dosageForm"] = m_dosage.group(1).strip()
        
        m_date = re.search(r'Test\s*date\s*:\s*([^\n\r]+)', sample_text, re.I)
        if m_date: sample["testDate"] = m_date.group(1).strip()
        
        # Processing Notes
        m_pnotes = re.search(r'Processing\s*Analyst\s*Notes?\s*:\s*(.*?)(?=(?:\n\s*Reading\s*Analyst\s*notes?|\n\s*Sample\s*History|\n\s*Conclusion|\n\s*Attn:|\n\s*---\s*\n|$))', sample_text, re.I | re.S)
        if m_pnotes:
            pnotes_str = m_pnotes.group(1).strip()
            sample["processingNotes"] = pnotes_str
            sample["analystNotes"] = pnotes_str
            
        # Reading Notes (inconclusive)
        m_rnotes = re.search(r'Reading\s*Analyst\s*notes?\s*:\s*(.*?)(?=(?:\n\s*Sample\s*History|\n\s*Conclusion|\n\s*Attn:|\n\s*---\s*\n|$))', sample_text, re.I | re.S)
        if m_rnotes:
            sample["readingNotes"] = m_rnotes.group(1).strip()
            
        # Sample History
        m_hist = re.search(r'Sample\s*History\s*:\s*(.*?)(?=(?:\n\s*Conclusion|\n\s*Attn:|\n\s*---\s*\n|$))', sample_text, re.I | re.S)
        if m_hist:
            sample["sampleHistory"] = m_hist.group(1).strip()
            
        # Conclusion
        m_conc = re.search(r'Conclusion\s*:\s*(.*?)(?=(?:\n\s*Attn:|\n\s*---\s*\n|$))', sample_text, re.I | re.S)
        if m_conc:
            conc_str = m_conc.group(1).strip()
            sample["conclusion"] = conc_str
            sample["customConclusion"] = conc_str
            
        # Extract volumeUsed from notes
        full_notes = sample["analystNotes"] or sample["processingNotes"]
        m_vol = re.search(r'(\d+(?:\.\d+)?\s*m[lL])\s+(?:of\s+(?:the\s+)?sample\s+was\s+(?:tested|filtered)|sample|was\s+filtered)', full_notes, re.I)
        if m_vol:
            sample["volumeUsed"] = m_vol.group(1).replace(" ", "")
        else:
            m_vol2 = re.search(r'analyst\s+filtered\s+a\s+(\d+(?:\.\d+)?\s*m[lL])\s+sample', full_notes, re.I)
            if m_vol2:
                sample["volumeUsed"] = m_vol2.group(1).replace(" ", "")
            else:
                sample["volumeUsed"] = "12mL"
                
        # Extract heating
        if "heated prior to filtration" in full_notes.lower():
            sample["isHeated"] = "heated"
        elif "heated fluid d" in full_notes.lower():
            sample["isHeated"] = "fluid_d"
        else:
            sample["isHeated"] = "none"
            
        # Extract Duplicate Lot / MS from Sample History
        hist = sample["sampleHistory"]
        m_dup = re.search(r'duplicate\s+Lot:\s*([^\s]+)\s+with\s+prior\s+(?:\"[^\"]+\"|[^\s]+)\s+result\s+under\s+(ETX-\d{6}-\d{4})', hist, re.I)
        if m_dup:
            sample["lotType"] = "duplicate"
            sample["origLot"] = m_dup.group(1)
            sample["origEtx"] = m_dup.group(2)
        else:
            sample["lotType"] = "new"
            sample["origLot"] = sample["lotNumber"]
            sample["origEtx"] = ""
            
        m_ms = re.search(r'Method\s+suitability\s+(ETX-\d{6}-\d{4})\s+with\s+(\d+(?:\.\d+)?\s*m[lL])', hist, re.I)
        if m_ms:
            sample["msStatus"] = "has"
            sample["msEtx"] = m_ms.group(1)
            sample["msVolume"] = m_ms.group(2).replace(" ", "")
        elif "no method suitability on file" in hist.lower():
            sample["msStatus"] = "none"
            sample["msEtx"] = ""
            sample["msVolume"] = ""
            
        # Match conclusion key
        c_lower = sample["conclusion"].lower()
        if "exceeding the specified" in c_lower and "cannot be confirmed as incompatible" in c_lower:
            sample["conclusionKey"] = "exceeded_duplicate"
        elif "exceeding the specified" in c_lower:
            sample["conclusionKey"] = "exceeded_new"
        elif "previously incompatible" in c_lower or "previously inconclusive" in c_lower:
            sample["conclusionKey"] = "duplicate"
        elif "new lot with no prior test results" in c_lower:
            sample["conclusionKey"] = "new"
        else:
            sample["conclusionKey"] = "custom"
            
        samples.append(sample)
        
    return samples

# Test Incompatible
email_inc = """
Hi @Eagle Client Care Team,

After review, please inform the client that the following sample has been found to yield incompatible results after Scan RDI testing:

AnazaoHealth Corporation - Tampa
ETX-260908-0123

Sample: Morphine Sulfate 10 mg/mL
Lot #: 20260908@1
Dosage form: Injection
Test date: 08SEP26

Processing Analyst Notes: 12mL of the sample was tested. The sample was heated prior to filtration. A visible layer of residue was observed on the membrane after filtration.

Sample History: Method suitability ETX-250801-0001 with 12mL per FIFU/Scan Filter Unit was specified.

Conclusion: As the sample is a new lot with no prior test results, additional sample vials may be required for lot compatibility verification.

Attn: @Elysse Nioupin, @Andrew Carrillo, @Mukyung Jang, @Ishita Sharma, @Sasha Allen.
"""

res = parse_completed_email(email_inc, "incompatible")
print("Incompatible parsed count:", len(res))
for k, v in res[0].items():
    print(f"  {k}: {repr(v)}")

# Test Inconclusive
email_inconc = """
Hi @Eagle Client Care Team,

After review, please inform the client that the following sample has been found to yield inconclusive results after Scan RDI testing:

Optimal Balance Pharmacy
ETX-260914-0655

Sample: Testosterone Cypionate 200 mg/mL
Dosage Form: Sesame Oil
Lot: 09142026@1
Test date: 14SEP26

Processing Analyst notes: 12mL of the sample was tested. The sample was heated prior to filtration. The sample and Fluid D filtered completely, but oil droplets remained on the membrane after filtration. Image has been uploaded to the Eagle Trax submission.

Reading Analyst notes: Scan RDI instrument displayed a high background baseline warning on the scan membrane.

Sample History: This sample is a duplicate Lot: 09142026@1 with prior inconclusive result under ETX-260814-0100. Method suitability ETX-260101-0001 with 10mL per FIFU/Scan Filter Unit was specified.

Conclusion: As this sample is a duplicate lot with previously incompatible result and may not be suitable for testing on the ScanRDI platform. Attn: @Elysse Nioupin, @Andrew Carrillo, @Mukyung Jang, @Ishita Sharma, @Sasha Allen.
"""
res2 = parse_completed_email(email_inconc, "inconclusive")
print("\nInconclusive parsed count:", len(res2))
for k, v in res2[0].items():
    print(f"  {k}: {repr(v)}")

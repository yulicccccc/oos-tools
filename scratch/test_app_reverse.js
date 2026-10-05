function parseCompletedEmail(input, type) {
    if (!input || typeof input !== 'string') return null;

    const lower = input.toLowerCase();
    const isCompletedEmail =
        (lower.includes('yield incompatible results after scan rdi testing') ||
         lower.includes('yield inconclusive results after scan rdi testing') ||
         lower.includes('processing analyst notes') ||
         lower.includes('processing analyst notes:')) &&
        (lower.includes('sample history:') || lower.includes('conclusion:') || lower.includes('conclusion')) &&
        /ETX-\d{6}-\d{4}/i.test(input);

    if (!isCompletedEmail) return null;

    let clean = input
        .replace(/<br\s*\/?>/gi, '\n')
        .replace(/<\/(?:p|div|tr|li|h[1-6])>/gi, '\n')
        .replace(/<[^>]+>/g, '')
        .replace(/&nbsp;/gi, ' ')
        .replace(/&amp;/gi, '&')
        .replace(/&lt;/gi, '<')
        .replace(/&gt;/gi, '>')
        .replace(/&quot;/gi, '"');

    const lines = clean.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);

    const etxIndices = [];
    for (let idx = 0; idx < lines.length; idx++) {
        const line = lines[idx];
        if (/\b(?:method suitability|prior|under|retest)\b/i.test(line)) continue;
        const m = line.match(/\b(ETX-\d{6}-\d{4})\b/i);
        if (m) {
            let isSubmission = false;
            if (/^(?:Submission\s*ID(?:\s*\(ETX\))?\s*:\s*)?ETX-\d{6}-\d{4}$/i.test(line)) {
                isSubmission = true;
            } else if (idx + 1 < lines.length && /^(?:Sample|Lot|Dosage|Test\s*date)/i.test(lines[idx + 1])) {
                isSubmission = true;
            } else if (idx + 2 < lines.length && /^(?:Sample|Lot|Dosage|Test\s*date)/i.test(lines[idx + 2])) {
                isSubmission = true;
            } else if (line.startsWith(m[1])) {
                isSubmission = true;
            }
            if (isSubmission) {
                etxIndices.push({ idx: idx, etx: m[1].toUpperCase() });
            }
        }
    }

    if (etxIndices.length === 0) return null;

    const samples = [];
    for (let i = 0; i < etxIndices.length; i++) {
        const item = etxIndices[i];
        let clientName = '';
        const ignore = ["after review", "hi @", "client care", "attention", "attn:", "subject:", "to:", "cc:", "scan rdi testing:"];
        for (let k = item.idx - 1; k >= Math.max(0, item.idx - 5); k--) {
            const cand = lines[k].trim();
            if (!cand) continue;
            const candLower = cand.toLowerCase();
            if (ignore.some(ig => candLower.includes(ig)) ||
                candLower.startsWith('conclusion:') ||
                candLower.startsWith('---') ||
                /ETX-\d{6}-\d{4}/i.test(cand)) {
                continue;
            }
            clientName = cand;
            break;
        }

        const startLine = item.idx;
        const endLine = (i + 1 < etxIndices.length) ? etxIndices[i + 1].idx : lines.length;
        const sampleText = lines.slice(startLine, endLine).join('\n');

        const sample = {
            clientName: clientName,
            submissionId: item.etx,
            sampleName: '',
            lotNumber: '',
            dosageForm: '',
            testDate: '',
            analystNotes: '',
            processingNotes: '',
            readingNotes: '',
            sampleHistory: '',
            conclusion: '',
            customConclusion: '',
            _isFromCompletedEmail: true
        };

        const mSample = sampleText.match(/(?:Sample(?:\s*name)?)\s*:\s*([^\n\r]+)/i);
        if (mSample) sample.sampleName = mSample[1].trim();

        const mLot = sampleText.match(/(?:Lot(?:\s*#|\s*number)?)\s*:\s*([^\n\r]+)/i);
        if (mLot) sample.lotNumber = mLot[1].trim();

        const mDosage = sampleText.match(/(?:Dosage\s*form)\s*:\s*([^\n\r]+)/i);
        if (mDosage) sample.dosageForm = mDosage[1].trim();

        const mDate = sampleText.match(/(?:Test\s*date)\s*:\s*([^\n\r]+)/i);
        if (mDate) sample.testDate = mDate[1].trim();

        // Processing notes
        const mPNotes = sampleText.match(/(?:Processing\s*Analyst\s*Notes?)\s*:\s*([\s\S]*?)(?=(?:\n\s*Reading\s*Analyst\s*notes?|\n\s*Sample\s*History|\n\s*Conclusion|\n\s*Attn:|\n\s*---\s*\n|$))/i);
        if (mPNotes) {
            const pStr = mPNotes[1].trim();
            sample.processingNotes = pStr;
            sample.analystNotes = pStr;
        }

        // Reading notes (inconclusive)
        const mRNotes = sampleText.match(/(?:Reading\s*Analyst\s*notes?)\s*:\s*([\s\S]*?)(?=(?:\n\s*Sample\s*History|\n\s*Conclusion|\n\s*Attn:|\n\s*---\s*\n|$))/i);
        if (mRNotes) {
            sample.readingNotes = mRNotes[1].trim();
        }

        // Sample history
        const mHist = sampleText.match(/(?:Sample\s*History)\s*:\s*([\s\S]*?)(?=(?:\n\s*Conclusion|\n\s*Attn:|\n\s*---\s*\n|$))/i);
        if (mHist) {
            sample.sampleHistory = mHist[1].trim();
        }

        // Conclusion
        const mConc = sampleText.match(/(?:Conclusion)\s*:\s*([\s\S]*?)(?=(?:Attn:|\n\s*---\s*\n|$))/i);
        if (mConc) {
            const cStr = mConc[1].trim();
            sample.customConclusion = cStr;
            sample.conclusion = cStr;
        }

        // Extract volumeUsed from notes
        const fullNotes = sample.analystNotes || sample.processingNotes;
        const mVol = fullNotes.match(/(\d+(?:\.\d+)?\s*m[lL])\s+(?:of\s+(?:the\s+)?sample\s+was\s+(?:tested|filtered)|sample|was\s+filtered)/i);
        if (mVol) {
            sample.volumeUsed = mVol[1].replace(/\s+/g, '');
        } else {
            const mVol2 = fullNotes.match(/analyst\s+filtered\s+a\s+(\d+(?:\.\d+)?\s*m[lL])\s+sample/i);
            if (mVol2) {
                sample.volumeUsed = mVol2[1].replace(/\s+/g, '');
            } else {
                sample.volumeUsed = '12mL';
            }
        }

        // Extract heating
        if (/heated\s+prior\s+to\s+filtration/i.test(fullNotes)) {
            sample.isHeated = 'heated';
        } else if (/heated\s+fluid\s+d/i.test(fullNotes)) {
            sample.isHeated = 'fluid_d';
        } else {
            sample.isHeated = 'none';
        }

        // Extract Duplicate Lot / MS from history
        const hist = sample.sampleHistory || '';
        const mDup = hist.match(/duplicate\s+Lot:\s*([^\s]+)\s+with\s+prior\s+(?:\"[^\"]+\"|[^\s]+)\s+result\s+under\s+(ETX-\d{6}-\d{4})/i);
        if (mDup) {
            sample.lotType = 'duplicate';
            sample.origLot = mDup[1];
            sample.origEtx = mDup[2];
        } else {
            sample.lotType = 'new';
            sample.origLot = sample.lotNumber || '';
            sample.origEtx = '';
        }

        const mMs = hist.match(/Method\s+suitability\s+(ETX-\d{6}-\d{4})\s+with\s+(\d+(?:\.\d+)?\s*m[lL])/i);
        if (mMs) {
            sample.msStatus = 'has';
            sample.msEtx = mMs[1];
            sample.msVolume = mMs[2].replace(/\s+/g, '');
        } else if (/no\s+method\s+suitability\s+on\s+file/i.test(hist)) {
            sample.msStatus = 'none';
            sample.msEtx = '';
            sample.msVolume = '';
        }

        // Map conclusion option key
        const cLower = (sample.conclusion || '').toLowerCase();
        if (cLower.includes('exceeding the specified') && (cLower.includes('cannot be confirmed') || cLower.includes('duplicate'))) {
            sample.conclusion = 'duplicate_exceeded';
        } else if (cLower.includes('exceeding the specified')) {
            sample.conclusion = 'new_exceeded';
        } else if (cLower.includes('previously incompatible') || cLower.includes('previously inconclusive') || cLower.includes('duplicate lot')) {
            sample.conclusion = 'duplicate';
        } else if (cLower.includes('new lot with no prior test results')) {
            sample.conclusion = 'new';
        } else {
            sample.conclusion = 'custom';
        }

        samples.push(sample);
    }

    return samples.length > 0 ? samples : null;
}

const test1 = `
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
`;

const parsed1 = parseCompletedEmail(test1, 'incompatible');
console.log('Test 1 (Incompatible):', parsed1 ? `Success! ${parsed1.length} sample(s)` : 'Failed');
if (parsed1) {
    console.log(JSON.stringify(parsed1[0], null, 2));
}

const test2 = `
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
`;

const parsed2 = parseCompletedEmail(test2, 'inconclusive');
console.log('\nTest 2 (Inconclusive):', parsed2 ? `Success! ${parsed2.length} sample(s)` : 'Failed');
if (parsed2) {
    console.log(JSON.stringify(parsed2[0], null, 2));
}

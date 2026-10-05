import sys
sys.path.append('scratch')
from test_reverse_parse import parse_completed_email

html_sample = '''
<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000;background-color:#fff;">
<p>Hi <a href="mailto:clientcare@eagleanalytical.com" style="color:#0000ee;text-decoration:underline;">@Eagle Client Care Team</a>,</p>
<p>After review, please inform the client that the following sample has been found to yield incompatible results after Scan RDI testing:</p>
<div style="margin-top: 20px;">
<p style="font-size:16pt;font-weight:bold;margin:0 0 4px 0;"><span style="background-color:yellow;">AnazaoHealth Corporation - Tampa</span></p>
<p style="margin:4px 0 12px 0;"><span style="background-color:yellow;">ETX-260908-0123</span></p>
<p style="margin:4px 0;">
<span style="font-weight:bold;text-decoration:underline;">Sample</span>: Morphine Sulfate 10 mg/mL<br>
<span style="font-weight:bold;text-decoration:underline;">Lot #</span>: 20260908@1<br>
<span style="font-weight:bold;">Dosage form</span>: Injection<br>
<span style="background-color:yellow;"><span style="font-weight:bold;text-decoration:underline;">Test date</span>: 08SEP26</span>
</p>
<p style="margin:14px 0 4px 0;">
<span style="font-weight:bold;text-decoration:underline;">Processing Analyst Notes:</span> 12mL of the sample was tested. The sample was heated prior to filtration. A visible layer of residue was observed on the membrane after filtration.
</p>
<p style="margin:14px 0 4px 0;">
<span style="font-weight:bold;text-decoration:underline;">Sample History:</span> Method suitability <span style="background-color:yellow;"><b>ETX-250801-0001</b> with 12mL per FIFU/Scan Filter Unit was specified.</span>
</p>
<p style="margin:14px 0 4px 0;">
<span style="font-weight:bold;text-decoration:underline;">Conclusion:</span> As the sample is a new lot with no prior test results, additional sample vials may be required for lot compatibility verification.
</p>
</div>
<p style="margin:16px 0 4px 0;">Attn: <a href="mailto:Elysse.Nioupin@eagleanalytical.com">@Elysse Nioupin</a>, <a href="mailto:Andrew.Carrillo@eagleanalytical.com">@Andrew Carrillo</a>.</p>
</div>
'''

res = parse_completed_email(html_sample, 'incompatible')
print('HTML parse count:', len(res) if res else 0)
if res:
    print('Client:', res[0]['clientName'])
    print('ETX:', res[0]['submissionId'])
    print('Sample:', res[0]['sampleName'])
    print('Lot:', res[0]['lotNumber'])
    print('Dosage:', res[0]['dosageForm'])
    print('Date:', res[0]['testDate'])
    print('Notes:', res[0]['analystNotes'])
    print('History:', res[0]['sampleHistory'])
    print('Conclusion:', res[0]['conclusion'])
    print('Vol:', res[0]['volumeUsed'])
    print('Heated:', res[0]['isHeated'])
    print('MS ETX:', res[0]['msEtx'])
    print('MS Vol:', res[0]['msVolume'])

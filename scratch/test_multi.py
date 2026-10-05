import re
from test_reverse_parse import parse_completed_email

multi_email = """
Hi @Eagle Client Care Team,

After review, please inform the client that the following samples have been found to yield incompatible results after Scan RDI testing:

Client One Medical
ETX-260908-0001

Sample: Morphine Sulfate 10 mg/mL
Lot #: 20260908@1
Dosage form: Injection
Test date: 08SEP26

Processing Analyst Notes: 12mL of the sample was tested. A visible layer of residue was observed on the membrane after filtration.

Sample History: Method suitability ETX-250801-0001 with 12mL per FIFU/Scan Filter Unit was specified.

Conclusion: As the sample is a new lot with no prior test results, additional sample vials may be required for lot compatibility verification.

Client Two Healthcare
ETX-260908-0002

Sample: Fentanyl Citrate 50 mcg/mL
Lot #: 20260908@2
Dosage form: Solution
Test date: 08SEP26

Processing Analyst Notes: 25mL of the sample was tested. The sample was heated prior to filtration. Residue remained.

Sample History: No method suitability on file.

Conclusion: As the sample is a new lot with no prior test results, additional sample vials may be required for lot compatibility verification.

Attn: @Elysse Nioupin, @Andrew Carrillo, @Mukyung Jang, @Ishita Sharma, @Sasha Allen.
"""

res = parse_completed_email(multi_email, "incompatible")
print("Multi-sample count:", len(res))
for i, s in enumerate(res):
    print(f"Sample {i+1}: ETX={s['submissionId']}, Client={s['clientName']}, SampleName={s['sampleName']}, Vol={s['volumeUsed']}, Heated={s['isHeated']}, MS={s['msStatus']}")

import re

mapping = {
    "KA": "Kathleen Aruta", "DH": "Domiasha Harrison", "GL": "Guanchen Li", "DS": "Devanshi Shah",
    "QC": "Qiyue Chen", "HS": "Halaina Smith", "MJ": "Mukyung Jang", "AS": "Alex Saravia",
    "CSG": "Clea S. Garza", "CGS": "Clea S. Garza", "RS": "Robin Seymour", "CCD": "Cuong Du", "VV": "Varsha Subramanian",
    "KS": "Karla Silva", "GS": "Gabrielle Surber", "PG": "Pagan Gary", "DT": "Debrework Tassew",
    "GA": "Gerald Anyangwe", "MRB": "Muralidhar Bythatagari", "TK": "Tamiru Kotisso", "OA": "Olugbenga Ajayi",
    "RE": "Rey Estrada", "AOD": "Ayomide Odugbesi", "EN": "Elysse Nioupin", "SU": "Sonal Uprety", 
    "AC": "Andrew Carrillo", "KC": "Kira C", "MC": "Maraya Chukwumerije",
    "AA": "America Alanis", "ALA": "America Alanis"
}

test_plates = [
    "ScanC/O CGS E001309 S1 11MAY2026",
    "ScanCO HS BSC1309 S3 17FEB2026",
    "Scan GA 1311 SettR 23FEB2026",
    "Scan PG BSC1310 Sett2 26FEB2026",
    "Scan DH BSC1313 SettL 26FEB2026",
    "Scan VV BSC1309 S1 27FEB2026",
    "Sterility GS BSC1314 Sett2 05FEB2026",
    "EM OA 114A cart 10FEB2026"
]

def extract_analyst_from_plate(p_name):
    # Match pattern after Scan, ScanC/O, ScanCO, Sterility, EM
    match = re.search(r"(?:ScanC/O|ScanCO|Scan|Sterility|EM)\s+([A-Z]{2,3})\b", p_name, re.IGNORECASE)
    if match:
        init = match.group(1).upper()
        if init in mapping:
            return init, mapping[init]
    return None, None

for p in test_plates:
    init, name = extract_analyst_from_plate(p)
    print(f"Plate: '{p}' => Initial: '{init}', Full Name: '{name}'")

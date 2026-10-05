import easyocr, os

reader = easyocr.Reader(['en'], gpu=False)

for i in range(1, 25):
    img_path = f"scratch/scans/193217/p{i}.png"
    # Read text
    res = reader.readtext(img_path, detail=0)
    print(f"\n=================== PAGE {i} ===================")
    # Print lines that contain dates or BSC or analysts or findings
    for r in res:
        # Check if line has any date, bsc, initial, result
        if any(k in r.lower() for k in ['114', '115', '1316', '1798', 'gs', 'ala', 'ccd', 'smo', '01sep', '02sep', '03sep', '07sep', '08sep', '09sep', '10sep', '15sep', '31aug', 'cfu', 'pass', 'growth', 'cleanroom']):
            print("  ", r)

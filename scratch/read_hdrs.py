import easyocr, glob, os

reader = easyocr.Reader(['en'], gpu=False)

for f in sorted(glob.glob('scratch/scans/*_hdr.png')):
    res = reader.readtext(f, detail=0)
    print(f"{os.path.basename(f)}: {' | '.join(res)}")

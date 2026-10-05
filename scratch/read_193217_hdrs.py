import easyocr, glob, os

reader = easyocr.Reader(['en'], gpu=False)

for i in range(1, 25):
    f = f"scratch/scans/193217/p{i}_hdr.png"
    res = reader.readtext(f, detail=0)
    print(f"Page {i:2d}: {' | '.join(res[:6])}")

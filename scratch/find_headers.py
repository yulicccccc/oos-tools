import glob, os
import easyocr, cv2

reader = easyocr.Reader(['en'], gpu=False)

for img_path in sorted(glob.glob('scratch/scans/*.png')):
    img = cv2.imread(img_path)
    h, w, _ = img.shape
    top = img[0:int(h*0.25), 0:w]
    res = reader.readtext(top, detail=0)
    joined = ' | '.join(res[:6])
    print(f'{os.path.basename(img_path)}: {joined}')

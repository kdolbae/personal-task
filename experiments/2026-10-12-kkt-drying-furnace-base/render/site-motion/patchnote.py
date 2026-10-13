import sys,glob,os
from multiprocessing import Pool
from PIL import Image,ImageDraw,ImageFont
F=ImageFont.truetype('/root/.fonts/Pretendard-Regular.otf',21)
def fix(p):
    im=Image.open(p);d=ImageDraw.Draw(im)
    d.rectangle((690,1024,1919,1079),fill=(13,25,34))
    d.text((714,1058),'크레인 위치·줄걸이는 개념 표현',font=F,fill=(233,239,242),anchor='ls')
    im.save(p,quality=93)
if __name__=='__main__':
    fs=[f for d in sys.argv[1:] for f in sorted(glob.glob(d+'/f_*.jpg'))]
    with Pool(4) as pl: pl.map(fix,fs,chunksize=50)
    print(len(fs))

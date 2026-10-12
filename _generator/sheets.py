import sys,glob
from PIL import Image, ImageDraw
d=sys.argv[1]
for pre in ("r1","r2"):
    fs=sorted(glob.glob(f"{d}/{pre}-*.png"))
    for k in range(0,len(fs),24):
        ims=[Image.open(f) for f in fs[k:k+24]]
        w,h=ims[0].size; cols=8; rows=(len(ims)+cols-1)//cols
        sheet=Image.new("RGB",(cols*(w+6),rows*(h+20)),"#888")
        for i,im in enumerate(ims):
            x=(i%cols)*(w+6); y=(i//cols)*(h+20)
            sheet.paste(im,(x,y+16)); ImageDraw.Draw(sheet).text((x+2,y+2),str(k+i+1),fill="white")
        sheet.save(f"{d}/sheet_{pre}_{k//24+1}.png")

from pathlib import Path
from PIL import Image,ImageOps,ImageDraw
import sys
folder=Path(sys.argv[1])
files=sorted(folder.glob('page*.png'))
for start in range(0,len(files),8):
    thumb_w=600; thumb_h=390
    sheet=Image.new('RGB',(thumb_w*2,thumb_h*4),'#dddddd')
    draw=ImageDraw.Draw(sheet)
    for i,p in enumerate(files[start:start+8]):
        im=Image.open(p).convert('RGB'); im.thumbnail((thumb_w-16,thumb_h-28))
        x=(i%2)*thumb_w+(thumb_w-im.width)//2; y=(i//2)*thumb_h+24
        sheet.paste(im,(x,y)); draw.text(((i%2)*thumb_w+10,(i//2)*thumb_h+5),p.name,fill='black')
    sheet.save(folder/f'montage-{start+1:02d}.jpg')

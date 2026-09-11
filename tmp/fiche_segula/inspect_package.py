from docx import Document
from zipfile import ZipFile
from pathlib import Path
from PIL import Image,ImageDraw
from lxml import etree
import json
p=Path('presentation/rapport/Fiche_Synthese_SEGULA_Capteur_Embarque_VAC_2.docx')
d=Document(p)
for n in ['Normal','Heading 1','Heading 2','Heading 3','Caption','Titre pdg','Titre pdg 1']:
    s=d.styles[n]; f=s.font; pf=s.paragraph_format
    print(n,f.name,f.size.pt if f.size else None,'bold',f.bold,'space',pf.space_before.pt if pf.space_before else None,pf.space_after.pt if pf.space_after else None)
with ZipFile(p) as z:
    print('ZIP integrity',z.testzip()); print('backslashes',[n for n in z.namelist() if '\\' in n])
    print('XML parse issues')
    for n in z.namelist():
        if n.endswith('.xml') or n.endswith('.rels'):
            try: etree.fromstring(z.read(n))
            except Exception as e: print(n,e)
    for n in ['word/_rels/document.xml.rels','word/header1.xml','word/footer1.xml']:
        if n in z.namelist(): (Path('tmp/fiche_segula')/Path(n).name).write_bytes(z.read(n))
mapping=[]
for i,p in enumerate(d.paragraphs):
    for blip in p._p.xpath('.//a:blip'):
        rid=blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
        part=d.part.related_parts[rid]
        path=Path('tmp/fiche_segula')/Path(str(part.partname)).name
        mapping.append((i,str(path),rid))
print('images',mapping)
im=Image.new('RGB',(1200,((len(mapping)+2)//3)*330),'#dddddd'); dr=ImageDraw.Draw(im)
for idx,(p,path,rid) in enumerate(mapping):
    pic=Image.open(path).convert('RGB'); pic.thumbnail((384,294)); x=idx%3*400+8;y=idx//3*330+28
    im.paste(pic,(x,y));dr.text((x,y-20),f'p{p} {rid} {Path(path).name}',fill='black')
im.save('tmp/fiche_segula/original_figures.jpg')

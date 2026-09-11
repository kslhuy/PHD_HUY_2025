from pathlib import Path
from pypdf import PdfReader
from docx import Document
import hashlib
root=Path.cwd();work=root/'tmp/fiche_segula'
r=PdfReader(work/'final_render/fiche.pdf')
out=[]
for i,page in enumerate(r.pages):
    t=page.extract_text()
    out.append(f'=== PAGE {i+1} ===\n{t}')
    ls=[l for l in t.splitlines() if l.strip()]
    print(i+1,len(t),'START', ' | '.join(ls[:3])[:160], 'END',' | '.join(ls[-3:])[:160])
(work/'final_text.txt').write_text('\n'.join(out),encoding='utf-8')
p=root/'presentation/rapport/Fiche_Synthese_SEGULA_Capteur_Embarque_VAC_3_enrichie.docx'
d=Document(p)
print('Images',len(d.inline_shapes),'tables',len(d.tables),'Sections',len(d.sections),'File size',p.stat().st_size)
print('Original hash',hashlib.sha256((root/'presentation/rapport/Fiche_Synthese_SEGULA_Capteur_Embarque_VAC_2.docx').read_bytes()).hexdigest())

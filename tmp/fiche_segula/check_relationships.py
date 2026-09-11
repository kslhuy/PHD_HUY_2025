from zipfile import ZipFile
from lxml import etree as E
import posixpath
z=ZipFile('presentation/rapport/Fiche_Synthese_SEGULA_Capteur_Embarque_VAC_2.docx')
ns={'r':'http://schemas.openxmlformats.org/package/2006/relationships'}
for n in z.namelist():
    if not n.endswith('.rels'): continue
    for rel in E.fromstring(z.read(n)):
        if rel.get('TargetMode')=='External':continue
        base=posixpath.dirname(posixpath.dirname(n)) if n!='_rels/.rels' else ''
        target=posixpath.normpath(posixpath.join(base,rel.get('Target'))).lstrip('/')
        if target not in z.namelist(): print('MISSING',n,rel.attrib,target)
for name in z.namelist():
    if not name.endswith('.xml'):continue
    root=E.fromstring(z.read(name))
    for el in root.iter():
        ign=el.get('{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable','')
        for prefix in ign.split():
            if prefix not in el.nsmap: print('BAD MC PREFIX',name,prefix)
print('root tags', [E.QName(el).localname for el in E.fromstring(z.read('word/document.xml'))[0]])

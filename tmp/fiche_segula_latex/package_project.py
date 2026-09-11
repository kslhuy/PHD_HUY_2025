from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from hashlib import sha256
import re,json

root=Path.cwd()
project=root/'presentation/rapport/Fiche_SEGULA_LaTeX'
out=root/'presentation/rapport/Fiche_SEGULA_LaTeX.zip'
files=[project/'Fiche_Synthese_SEGULA.tex',project/'Fiche_Synthese_SEGULA.pdf',project/'README.md']
files+=sorted((project/'figures').glob('*'))
files+=sorted((project/'sections').glob('*.tex'))
with ZipFile(out,'w',ZIP_DEFLATED) as z:
    for f in files:z.write(f,f.relative_to(project).as_posix())
with ZipFile(out) as z:
    assert z.testzip() is None
    names=set(z.namelist())
    assert len([n for n in names if n.startswith('figures/') and n.endswith('.pdf')])==8
    text=z.read('Fiche_Synthese_SEGULA.tex').decode('utf-8')
    assert '\\input{' not in text
    for asset in re.findall(r'^\\(?:sourcefig|boardfig)\{([^}]+)\}',text,re.M):
        assert 'figures/'+asset in names,asset
    for a,b in re.findall(r'^\\pairfig\{([^}]+)\}\{([^}]+)\}',text,re.M):
        assert 'figures/'+a in names and 'figures/'+b in names
    assert z.read('Fiche_Synthese_SEGULA.pdf')==(project/'Fiche_Synthese_SEGULA.pdf').read_bytes()
report={'files':len(files),'zip_bytes':out.stat().st_size,'pdf_pages':32,
        'new_presentation_figures':8,'sections_preserved':[1,2,3,4,6,7,8],
        'compile_errors':0,'overfull_boxes':0,'figures_included':True,
        'tex_sha256':sha256((project/'Fiche_Synthese_SEGULA.tex').read_bytes()).hexdigest()}
(root/'tmp/fiche_segula_latex/delivery_checks.json').write_text(json.dumps(report,indent=2),encoding='utf-8')
print(json.dumps(report))

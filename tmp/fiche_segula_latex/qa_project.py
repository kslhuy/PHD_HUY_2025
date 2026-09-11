from pathlib import Path
from pypdf import PdfReader
from docx import Document
from hashlib import sha256
import re,json,sys
sys.stdout.reconfigure(encoding='utf-8')
root=Path.cwd();project=root/'presentation/rapport/Fiche_SEGULA_LaTeX';work=root/'tmp/fiche_segula_latex'
r=PdfReader(project/'Fiche_Synthese_SEGULA.pdf')
texts=[p.extract_text() for p in r.pages]
(work/'compiled_text.txt').write_text('\n'.join(f'=== PDF PAGE {i+1} ===\n{t}' for i,t in enumerate(texts)),encoding='utf-8')
for i,t in enumerate(texts):
    lines=[s for s in t.splitlines() if s.strip()]
    print(i+1,'chars',len(t),'START',' | '.join(lines[2:4])[:105],'END',' | '.join(lines[-3:])[:120])
main=(project/'Fiche_Synthese_SEGULA.tex').read_text(encoding='utf-8')
assert '\\input{' not in main
for name in ['01_04_sections_conservees','06_08_sections_conservees','05_travaux_realises']:
    assert (project/f'sections/{name}.tex').read_text(encoding='utf-8') in main
manifest=json.loads((work/'source_manifest.json').read_text(encoding='utf-8'))
assert sha256((root/'presentation/rapport/Fiche_Synthese_SEGULA_Capteur_Embarque_VAC_2.docx').read_bytes()).hexdigest()==manifest['original_sha256']
assert len(re.findall(r'^\\sourcefig\{',main,re.M))==8
for f in manifest['new_figures_source_pages']:
    assert len(PdfReader(project/f'figures/{f}.pdf').pages)==1
assert not re.search(r'Overfull|LaTeX Warning|Missing character|LaTeX Error', (project/'Fiche_Synthese_SEGULA.log').read_text(encoding='utf-8',errors='replace'))
print('Verified: complete source, 8 PDF figures, preserved sections, original unchanged, clean compile.')

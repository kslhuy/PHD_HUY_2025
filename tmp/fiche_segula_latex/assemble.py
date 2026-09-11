from pathlib import Path
import re
root=Path.cwd();project=root/'presentation/rapport/Fiche_SEGULA_LaTeX';work=root/'tmp/fiche_segula_latex'
main=project/'Fiche_Synthese_SEGULA.tex'
template=work/'main_template.tex'
if not template.exists():template.write_text(main.read_text(encoding='utf-8'),encoding='utf-8')
text=template.read_text(encoding='utf-8')
def insert(m):
    path=project/(m.group(1)+'.tex')
    return '% Début '+m.group(1)+'\n'+path.read_text(encoding='utf-8')+'\n% Fin '+m.group(1)
text=re.sub(r'\\input\{([^}]+)\}',insert,text)
main.write_text(text,encoding='utf-8')
print('Single complete TeX file:',len(text),'characters')

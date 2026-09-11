from pathlib import Path
from zipfile import ZipFile
from docx import Document
from pypdf import PdfReader,PdfWriter
import json,re,hashlib
root=Path.cwd();base=root/'presentation/rapport';project=base/'Fiche_SEGULA_LaTeX'
doc=Document(base/'Fiche_Synthese_SEGULA_Capteur_Embarque_VAC_2.docx')
p=doc.paragraphs
def esc(s):
    repl={'\\':r'\textbackslash{}','&':r'\&','%':r'\%','$':r'\$','#':r'\#','_':r'\_','{':r'\{','}':r'\}','~':r'\textasciitilde{}','^':r'\textasciicircum{}','\u00a0':'~','→':r'\(\rightarrow\)','²':r'\textsuperscript{2}'}
    return ''.join(repl.get(c,c) for c in s)
def paras(indices):
    out=[];listing=False
    for i in indices:
        par=p[i];s=par.style.name;text=par.text
        if not text.strip():continue
        if s!='List Paragraph' and listing:out.append('\\end{itemize}\n');listing=False
        if s=='List Paragraph':
            if not listing:out.append('\\begin{itemize}');listing=True
            out.append('\\item '+esc(text))
        elif s.startswith('Heading'):
            name={1:'section',2:'subsection',3:'subsubsection',4:'paragraph'}[int(s[-1])]
            out.append('\\'+name+'{'+esc(text)+'}\n')
        else:out.append(esc(text)+'\n')
    if listing:out.append('\\end{itemize}')
    return '\n'.join(out)+'\n'
front=paras(range(31,105))
(project/'sections/01_04_sections_conservees.tex').write_text('% Texte des sections 1 à 4 repris du DOCX original, sans ajout de fond.\n'+front,encoding='utf-8')
back=paras(range(170,174))
back+='\n\\clearpage\n\\begin{landscape}\n'+paras(range(182,184))
back+='\nLes annexes suivantes à remplir lors de la transmission pour relecture de la FS et ne doivent pas être diffusées en externe.\n\n'
tab=doc.tables[-1]
rows=[[c.text for c in row.cells] for row in tab.rows]
widths=[2.0,4.0,7.7,6.2,1.9,1.8]
fmt='|'+ '|'.join('>{\\raggedright\\arraybackslash}p{'+str(w)+'cm}' for w in widths)+'|'
back+='\\begingroup\\scriptsize\n\\setlength{\\tabcolsep}{3pt}\n\\begin{longtable}{'+fmt+'}\n\\hline\n'
back+='\\rowcolor{TableHead} '+ ' & '.join('\\textbf{'+esc(c)+'}' for c in rows[0])+r' \\ \hline'+'\n\\endfirsthead\n'
back+='\\rowcolor{TableHead} '+ ' & '.join('\\textbf{'+esc(c)+'}' for c in rows[0])+r' \\ \hline'+'\n\\endhead\n'
for row in rows[1:]:back+=' & '.join(esc(c).replace('\n','\\par ') for c in row)+r' \\ \hline'+'\n'
back+='\\end{longtable}\n\\endgroup\n\\end{landscape}\n'
(project/'sections/06_08_sections_conservees.tex').write_text('% Texte des sections 6 à 8 et cellules de l’annexe repris du DOCX original.\n'+back,encoding='utf-8')
metadata=doc.tables[0]
meta_parts=[];seen_cells=set()
for row in metadata.rows:
    for c in row.cells:
        if c._tc in seen_cells:continue
        seen_cells.add(c._tc)
        meta_parts.append(esc(c.text).replace('\n','\\par '))
meta='\n\n'.join(meta_parts)
(project/'sections/00_identification.tex').write_text(meta+'\n',encoding='utf-8')
mapping={'rId11':'logo_segula.png','rId82':'carte_dediee.jpg','rId83':'architecture_carte.png','rId84':'carte_stm32.jpg','rId85':'chaine_complete.jpg','rId86':'cycle_ekf.png','rId87':'ros_topics.jpg','rId88':'ros_gnss.jpg','rId89':'test_flux.jpg','rId90':'test_pertes.jpg','rId91':'vehicule_capteur.jpg','rId92':'ecran_capteur.jpg','rId93':'resultats_ekf.png','rId94':'platoon_1.jpg','rId95':'platoon_2.jpg'}
for rid,name in mapping.items():(project/'figures'/name).write_bytes(doc.part.related_parts[rid].blob)
r=PdfReader(base/'Final_PHD_these_2026.pdf')
newfigs={'trust_architecture':9,'trust_poids':25,'trust_simulation':27,'rollback':32,'rknet_principe':34,'rknet_chaine':36,'rknet_resultats':45,'rknet_masques_gains':46}
for name,num in newfigs.items():
    w=PdfWriter();w.add_page(r.pages[num-1]);w.write(project/'figures'/f'{name}.pdf')
manifest={'original_sha256':hashlib.sha256((base/'Fiche_Synthese_SEGULA_Capteur_Embarque_VAC_2.docx').read_bytes()).hexdigest(),'preserved_paragraphs_1_4':list(range(31,105)),'preserved_paragraphs_6_7':list(range(170,174)),'preserved_annex_table_cells':rows,'new_figures_source_pages':newfigs,'source_slide_offset':29}
(root/'tmp/fiche_segula_latex/source_manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding='utf-8')
print('Sections conserved; original figures',len(mapping),'new figures',len(newfigs))

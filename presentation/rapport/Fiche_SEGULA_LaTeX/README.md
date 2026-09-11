# Fiche SEGULA en LaTeX

Le projet contient la fiche complète et huit nouvelles figures issues du support `Final_PHD_these_2026.pdf`.

## Ouvrir et modifier

- `Fiche_Synthese_SEGULA.tex` : fichier principal complet, à compiler avec **XeLaTeX**.
- `Fiche_Synthese_SEGULA.pdf` : version compilée et contrôlée.
- `sections/05_travaux_realises.tex` : texte du chapitre 5 seul, pour faciliter sa reprise dans un autre document.
- `figures/` : illustrations de la fiche originale et huit pages extraites du support de présentation, conservées en PDF.

Le fichier principal est autonome pour le texte. Il requiert le dossier `figures/` placé à côté de lui. Les fichiers dans `sections/` servent de copies séparées ; après consolidation du fichier principal, leurs modifications ne sont pas automatiquement reportées dans celui-ci.

## Compiler

Dans ce dossier :

```text
xelatex -interaction=nonstopmode -halt-on-error Fiche_Synthese_SEGULA.tex
xelatex -interaction=nonstopmode -halt-on-error Fiche_Synthese_SEGULA.tex
```

Sur Overleaf, importer le ZIP, sélectionner `Fiche_Synthese_SEGULA.tex` comme document principal et **XeLaTeX** comme compilateur. Le document utilise Arial lorsqu'elle est disponible, avec une police de remplacement prévue dans le préambule.

## Périmètre éditorial

Les sections 1 à 4 et 6 à 8 sont transcrites depuis `Fiche_Synthese_SEGULA_Capteur_Embarque_VAC_2.docx`. Les ajouts et la réécriture technique se concentrent dans le chapitre 5. Les tableaux de l'annexe sont conservés, y compris leurs indications de remplissage. Cette conservation n'est pas une nouvelle vérification bibliographique des références de la fiche d'origine.

Le chapitre 5 développe l'architecture de la carte, l'EKF et le firmware, le Trust et les poids de consensus, le retour arrière, Robust KalmanNet, leurs résultats et leur intégration matérielle. Les résultats déjà présents dans les sources sont distingués des propositions de portage et des exemples de dimensionnement. Les exemples mémoire et temps de calcul ne sont pas des mesures du STM32.

## Figures nouvelles et traçabilité

Les numéros ci-dessous sont les pages du fichier PDF, et non les numéros de diapositive affichés.

| Figure | Page PDF | Diapositive affichée |
|---|---:|---:|
| Architecture du Trust | 9 | 38 |
| Trust et poids de consensus | 25 | 54 |
| Résultats de simulation du Trust | 27 | 56 |
| Retour arrière | 32 | 61 |
| Robust KalmanNet State Estimation | 34 | 63 |
| Chaîne de calcul RKNet | 36 | 65 |
| Trajectoire et erreurs RKNet | 45 | 74 |
| Masques et gains RKNet | 46 | 75 |

Les graphiques d'origine sont conservés. Aucun graphique de résultat nouveau ni aucune série de mesures fictive n'a été ajouté.

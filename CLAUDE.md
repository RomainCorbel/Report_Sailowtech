# CLAUDE.md

Conventions pour toute session Claude Code travaillant dans ce dépôt.

## Ce qu'est ce dépôt

Rapport de projet de semestre ME-401 rédigé en LaTeX, décrivant la conception
d'une station météo embarquée pour les expéditions à la voile (projet Sailowtech).
Réalisé par Guilhem Destriau et Romain Corbel sous la supervision d'Eric Boillat
(LMTM — Thermomechanical Metallurgy Laboratory, EPFL).

## Structure des fichiers

- `main.tex` — préambule et `\input{sections/...}` de chaque section. Ne pas
  mettre de prose ici ; conserver comme coquille mince.
- `sections/*.tex` — un fichier par section du rapport.
- `references.bib` — bibliographie.
- `images/` — figures et images incluses dans le rapport.

## Style rédactionnel

- Français : orthographe et typographie françaises.
- Première personne du pluriel ("nous avons conçu") pour les descriptions de
  travaux effectués. Infinitif pour les étapes.
- Présent pour les descriptions ; infinitif ou impératif pour les procédures.
- Termes techniques introduits en italique à la première occurrence, romains
  ensuite.
- Phrases courtes. Préférer des verbes concrets aux noms abstraits.

## Conventions LaTeX

- `\cite{}` pour toutes les références. Chaque clé doit exister dans
  `references.bib` ; ne jamais inventer de clé.
- Utiliser `booktabs` (`\toprule`, `\midrule`, `\bottomrule`) ; pas de filets
  verticaux dans les tableaux.
- Préférer `\texttt{}` pour le code en ligne et `lstlisting` pour les blocs.
- Les figures doivent avoir une `\caption` et être référencées dans le texte
  avec `\ref{}`.
- Utiliser `\label{sec:...}` pour les sections, `\label{fig:...}` pour les
  figures et `\label{tab:...}` pour les tableaux.

## Ne pas committer

Artefacts de compilation : `*.aux`, `*.log`, `*.out`, `*.toc`, `*.bbl`,
`*.blg`, `*.fls`, `*.fdb_latexmk`, `*.synctex.gz` et les PDF compilés.
Couverts par `.gitignore`.

## Style de commit

- Mode impératif, ≤ 72 caractères pour la ligne d'objet.
- Décrire le *pourquoi*, pas le *quoi* (le diff montre le quoi).
- Un changement logique par commit.

## Lors des modifications

- Faire la modification minimale qui satisfait la demande.
- Pour les modifications globales, produire un plan d'abord ; attendre
  l'approbation avant d'écrire.
- Quand on ajoute une nouvelle section, l'ajouter à la fois comme fichier
  `sections/<nom>.tex` et comme `\input{sections/<nom>}` dans `main.tex`,
  dans l'ordre de lecture voulu.

## CAB PowerPoint Generator

Génère un support PowerPoint pour le Comité des Changements (CAB) à partir d’un jeu de données CSV/Excel et d’un modèle PPTX.

Ce que le script produit

- Timeline S+1: rectangles positionnés entre dates/horaires de début et fin, colorés par Type.
- Slides de détail: une par changement inclus dans la S+1 (titre avec RFC + résumé, tableau de champs).
- Optionnel S-1: 1) Camembert par code de fermeture, 2) Liste des changements non « Succès » de la semaine précédente.
- Optionnel S (semaine en cours): une timeline de la semaine courante en fin de support.

Caractéristiques visuelles

- Couleurs par type (timeline): Urgent → orange, Normal → bleu, Agile → vert (modifiable via `timeline_colors` dans le fichier de configuration JSON).
- Le numéro de RFC est un lien cliquable: `https://outils.change.fr/change=<rfc>`.
- Texte des boîtes de la timeline:
  - Ligne 1: « RFC – Résumé » (le Résumé est masqué si la boîte est trop étroite).
  - Ligne 2: « Date/heure début planifié » (format `dd/mm/YYYY HH:MM`).
  - Ligne 3: « Date/heure fin planifiée » (format `dd/mm/YYYY HH:MM`).

### Installation

Prérequis Python 3.10+

- `pip install -r requirements.txt`

### Données attendues

Colonnes requises (casse/accents respectés):

- `Numéro`
- `Type`
- `Etat`
- `Date de début planifiée`
- `Date de fin planifiée`

Le parsing des dates est robuste (formats `dd/mm/yy`, `dd/mm/YYYY`, `YYYY-MM-DD`, avec/ sans heures `HH:MM[:SS]`).

### Modèle PPTX

- Le script peut utiliser le placeholder « Titre » des layouts lorsque présent; sinon, un textbox de titre est ajouté en haut du slide.
- Par défaut, la timeline S+1 est dessinée sur la première diapositive du modèle. Vous pouvez choisir un autre layout via un flag (voir ci‑dessous).

## Utilisation

Commande type:

```
python generate_cab_pptx.py \
  --data cab_changes.xlsx \
  --template template_change_SPLUS1.pptx \
  --out change_generated.pptx \
  --ref-date 2025-09-09 \
  --sminus1-pie --current-week
```

### Options de ligne de commande

Obligatoires:

- `--data PATH` — fichier CSV/Excel des changements
- `--template PATH` — modèle PPTX
- `--out PATH` — fichier PPTX de sortie

Générales:

- `--ref-date YYYY-MM-DD` — date de référence (par défaut: aujourd’hui)
- `--encoding ENC` — encodage CSV forcé (ex: `cp1252`, `latin1`, `utf-8-sig`)
- `--sep SEP` — séparateur CSV forcé (ex: `;`, `,`, `\t`)
- `--list-layouts` — affiche les layouts du modèle et quitte
- `--include-tags TAGS` — filtre les changements en ne gardant que ceux dont la colonne `Balises` contient au moins une des balises listées (séparées par des virgules). Exemple: `RED_TRUC-TEL,GRE_BIDULE-PDT`.

Mise en page:

- `--detail-layout-index N` — layout pour les slides de détail
- `--splus1-layout-index N` — layout de la timeline S+1 (sinon: première slide du modèle)
- `--sminus1-layout-index N` — layout des slides S-1
- `--current-week-layout-index N` — layout de la slide « semaine en cours »
- `--assignee-layout-index N` — layout de la slide « répartition par affecté » (bar chart)

Slides optionnelles:

- `--sminus1-pie` — ajoute les deux slides S-1 (camembert + liste non « Succès »)
- `--current-week` — ajoute la timeline de la semaine en cours en fin de support

### Configuration JSON

- `timeline_colors`: dictionnaire de codes HTML `#RRGGBB` permettant de surcharger les couleurs des boîtes de la timeline par type (`"urgent"`, `"normal"`, `"agile"`). Exemple:

```
{
  "timeline_colors": {
    "urgent": "#FF8800",
    "normal": "#1F77B4",
    "agile": "#2CA02C"
  }
}
```

Les clés sont insensibles à la casse/accents; toute valeur invalide est ignorée avec un avertissement et les couleurs par défaut sont conservées.

### Exemples

- Générer S+1 uniquement (timeline + détails):

```
python generate_cab_pptx.py --data cab.xlsx --template template.pptx --out out.pptx
```

- Forcer des layouts spécifiques:

```
python generate_cab_pptx.py \
  --data cab.xlsx --template template.pptx --out out.pptx \
  --detail-layout-index 2 --splus1-layout-index 1 \
  --sminus1-pie --sminus1-layout-index 3 \
  --current-week --current-week-layout-index 4 \
  --assignee-layout-index 5
```

## Notes

- Le script détecte et utilise automatiquement le placeholder « Titre » des layouts; sinon, un titre est ajouté en haut.
- Pour diagnostiquer les layouts, utilisez `--list-layouts`.
- Les dates/heures des boîtes s’affichent toujours; le Résumé est masqué si la boîte est trop étroite.

## Développement

Code organisé par responsabilités:

- `data_loader.py` — chargement CSV/Excel + parsing robuste des dates
- `periods.py` — bornes de semaines S, S+1, S-1
- `layouts.py` — sélection/liste des layouts
- `render/` — rendu des diapositives (timeline, détails, utilitaires)
- `services.py` — orchestration (préparation des données, filtrage, génération de base)
- `generate_cab_pptx.py` — CLI principale

Contributions bienvenues via PRs.

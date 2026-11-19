# Spécification JSON pour la génération `.docx`

Le script `scripts/build_docx.py` produit un document Word conforme au genre demandé à partir d'un fichier JSON structuré, applique la typographie française et purge automatiquement les métadonnées XML ainsi que les paramètres d'URL.

## 1. Paramètres automatiques par genre

| Genre (`genre`) | Police | Corps / Interligne | Marges (cm) | Page de garde | Sommaire | En-tête et pagination | Palette par défaut |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| `memoire` | Times New Roman | 12 pt / 1,5 | 3 gauche, 2,5 autres | Oui | Table des matières | Oui (masqués page 1) | `monochrome` |
| `rapport` | Calibri | 11 pt / 1,15 | 2,5 partout | Oui | Sommaire | Oui (masqués page 1) | `bleu-nuit` |
| `expose` | Calibri | 11 pt / 1,15 | 2,5 partout | Non (bloc haut p. 1) | Non | Oui (masqués page 1) | `bleu-nuit` |
| `note` | Calibri | 11 pt / 1,15 | 2,5 partout | Non (bloc émetteur/objet) | Non | Non | `bleu-nuit` |

## 2. Champs racine du JSON

| Clé | Type | Description |
| :--- | :--- | :--- |
| `genre` | chaîne | `memoire`, `rapport`, `expose` ou `note` (défaut `rapport`). |
| `title` | chaîne (obligatoire) | Titre complet du document. |
| `short_title` | chaîne | Titre abrégé affiché dans l'en-tête courant. |
| `subtitle` | chaîne | Sous-titre éventuel sous le titre principal. |
| `document_type_label` | chaîne | Mention discrète au-dessus du titre (ex. *Audit opérationnel*). |
| `organization` | chaîne | Établissement, entreprise ou service. |
| `author` | chaîne | Nom réel de l'auteur ou du groupe (renseigné aussi dans `core.xml`). |
| `recipient` | chaîne | Destinataire (direction, client, jury). |
| `supervisor` | chaîne | Encadrant ou tuteur (pour un mémoire ou rapport de stage). |
| `program` | chaîne | Filière ou formation. |
| `date` | chaîne | Date du document. |
| `academic_year` | chaîne | Année académique (ex. *2025-2026*). |
| `reference` | chaîne | Référence interne pour une note ou un rapport. |
| `subject` | chaîne | Objet explicite pour le genre `note`. |
| `logo` | chaîne | Chemin local vers un logo (PNG ou JPEG) placé sur la page de garde. |
| `font` | chaîne | Doit appartenir aux polices natives : `Times New Roman`, `Garamond`, `Georgia`, `Calibri`, `Aptos`, `Cambria`, `Arial`. |
| `palette` | chaîne | `monochrome`, `bleu-nuit`, `ardoise`, `vert-foret`, `bordeaux`, ou hexadécimal `#RRGGBB` sobre. |
| `lang` | chaîne | `fr` (défaut) ou `en`. |
| `cover_page` | booléen | Force l'activation ou la désactivation de la page de garde. |
| `toc` | booléen | Force l'activation ou la désactivation du champ Table des matières. |
| `toc_title` | chaîne | Libellé du sommaire (*Sommaire* ou *Table des matières*). |
| `preliminaries` | tableau | Sections liminaires pour un mémoire (remerciements, résumé, sigles), numérotées en chiffres romains. |
| `sections` | tableau | Liste des titres et blocs de contenu du corps principal. |
| `bibliography_title` | chaîne | Titre de la section finale de sources (défaut *Bibliographie et sources*). |
| `bibliography` | tableau | Chaînes simples ou objets `{ "citation": "...", "url": "https://..." }`. |

## 3. Structure des sections et des blocs

Chaque élément de `sections` contient `level` (`1`, `2` ou `3`), `title`, un booléen optionnel `page_break_before`, et `blocks` :

1. **Paragraphe simple** : `{ "type": "paragraph", "text": "..." }`
2. **Paragraphe avec liens ou mises en forme** : `{ "type": "paragraph", "runs": [ { "text": "Voir " }, { "text": "le décret", "url": "https://www.legifrance.gouv.fr/..." } ] }`
3. **Liste à puces ou numérotée** : `{ "type": "bullets", "items": ["Premier élément.", "Deuxième élément."] }` (ou `"type": "numbered"`).
4. **Citation en retrait** : `{ "type": "quote", "text": "..." }`
5. **Tableau** : `{ "type": "table", "caption": "Tableau 1 : ...", "headers": [...], "rows": [[...]], "alignments": ["left", "right"], "col_widths_cm": [8.0, 4.0], "source": "..." }`
6. **Image** : `{ "type": "image", "path": "graphique.png", "caption": "Figure 1 : ...", "width_cm": 13.5 }`
7. **Saut de page** : `{ "type": "page_break" }`

## 4. Commandes usuelles

```bash
python scripts/build_docx.py spec.json -o "Rapport - Sujet.docx"
python scripts/audit_doc.py "Rapport - Sujet.docx" --strict
uv run scripts/render_preview.py "Rapport - Sujet.docx"
```

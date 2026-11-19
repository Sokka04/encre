# Encre

[English](#english) · [Français](#français)

---

<a id="english"></a>

## English

**Write like a human, not an LLM.** Encre is an editorial skill and `.docx` audit suite for **Claude Code** and **Antigravity**.

It replaces default AI habits (em-dash clutter, padded introductions, repetitive bullet blocks, invented citations, neon purple headings, `output_v1_final.docx` filenames, tracking URLs, and library fingerprints in XML metadata) with direct prose, real sources, sober native typography, and a blocking verification pipeline before delivery.

### Before / After

```text
Before (default LLM draft):
In today's fast-paced world — and it is worth noting — digital
transformation is not just a tool, it is a holistic paradigm shift 🚀.
Moreover, it is crucial to underscore that this robust ecosystem leverages
synergistic levers (source: https://example.org/report?utm_source=chatgpt).

After (with Encre):
Since March 2024, the twelve regional branches have digitized invoice intake.
Processing time dropped from four days to eleven hours across 18,400 files,
mainly by removing duplicate manual checks at dispatch.
```

### How the 6-phase protocol works

| Phase | What Encre enforces | Tools and references |
| :--- | :--- | :--- |
| **0. Clarify first** | Asks once (3 to 6 concise questions) for author, organization, recipient, exact title, local logo file, or internal data instead of guessing or leaving `[INSERT LOGO]` placeholders. | `SKILL.md` |
| **1. Genre-aware structure** | Adapts layout to the document type: full academic apparatus for a thesis (`memoire`), concise cover + TOC for a business report (`rapport`), clean top header without bloat for a short brief (`expose`), or direct memo header (`note`). | [docx-spec.md](./references/docx-spec.md) |
| **2. Human prose rules** | Strips em-dashes in prose, bold-lead bullet spam, `"not just X, it's Y"` clichés, stacked paragraph connectors, self-justifying filler, and inflated AI vocabulary in French and English. | [style-anti-ia.md](./references/style-anti-ia.md) |
| **3. Verified sources & clean URLs** | Prioritizes primary/official sources, forbids fabricated references, and strips tracking and assistant tokens (`utm_source`, `fbclid`, `ref=`, `=chatgpt`, `=claudecode`, `=gemini`) from visible text and hidden Word hyperlinks. | [citations-sources.md](./references/citations-sources.md) |
| **4. Publication-grade `.docx`** | Generates Word files with native fonts (`Times New Roman`, `Calibri`, `Aptos`, `Cambria`, `Garamond`, `Georgia`, `Arial`), sober palettes, real TOC fields, `w:tblHeader` / `w:cantSplit` tables, and sanitized XML (`core.xml`, `app.xml`, `custom.xml`, `comments.xml`, `word/_rels/*.rels`). | [build_docx.py](./scripts/build_docx.py), [sanitize_docx.py](./scripts/sanitize_docx.py) |
| **5. Blocking audit & visual check** | Runs deterministic checks across body prose, tables, footnotes, headers/footers, comments, and XML metadata, then renders PNG previews of every page via Word or LibreOffice for visual inspection. | [audit_doc.py](./scripts/audit_doc.py), [render_preview.py](./scripts/render_preview.py) |

### Installation

#### Claude Code

macOS / Linux:

```bash
git clone https://github.com/Sokka04/encre.git ~/.claude/skills/encre
```

Windows (PowerShell):

```powershell
git clone https://github.com/Sokka04/encre.git "$env:USERPROFILE\.claude\skills\encre"
```

#### Antigravity

macOS / Linux:

```bash
git clone https://github.com/Sokka04/encre.git ~/.gemini/config/skills/encre
```

Windows (PowerShell):

```powershell
git clone https://github.com/Sokka04/encre.git "$env:USERPROFILE\.gemini\config\skills\encre"
```

Once installed, invoke `/encre` or simply ask the agent to write, rewrite, humanize, or build a `.docx` report, thesis, brief, or memo.

### Prerequisites and CLI commands

- Python 3.11+ (`sanitize_docx.py` and `audit_doc.py` rely on the Python standard library only)
- `uv` (recommended; automatically installs `python-docx`, `pillow`, and `pymupdf` declared in `build_docx.py` and `render_preview.py`)
- Microsoft Word on Windows or LibreOffice (`soffice`) so `render_preview.py` can export `.docx` to PDF, update the Table of Contents, and render PNG pages

```bash
# Build a clean .docx from a structured JSON spec
uv run scripts/build_docx.py examples/rapport-exemple.json -o "Rapport - Livraisons Tours.docx"

# Sanitize XML metadata, comments, and tracking URLs in an existing .docx
python scripts/sanitize_docx.py "Existing Report.docx" --in-place --author "Claire Martin" --title "Annual Audit"

# Run the strict audit (prose, tables, comments.xml, URLs, AI signatures, typography)
python scripts/audit_doc.py "Rapport - Livraisons Tours.docx" --strict

# Render every page to PNG for visual inspection
uv run scripts/render_preview.py "Rapport - Livraisons Tours.docx"
```

### Guardrails and Responsible Use

#### 1. Academic and professional integrity
Encre improves clarity, structure, typography, and formatting hygiene. It is not designed for plagiarism, fabricating work, or bypassing university, school, or employer AI-disclosure rules. You remain solely responsible for the accuracy and authorship of what you submit.

#### 2. Fact and source verification
Automated scripts catch malformed links, tracking tokens (`=chatgpt`, `=claudecode`, `=gemini`, `utm_source`), and stylistic clichés, not factual truth. Always verify statistics, legal citations, quotes, and bibliography entries yourself before publishing or submitting.

#### 3. What XML sanitization actually does
`sanitize_docx.py` replaces technical generator tags (such as `python-docx` or `Pandoc`), strips tracking query parameters, and cleans AI markers in document properties and `word/comments.xml` so the file opens cleanly like a standard Word document. Clean metadata never compensates for weak thinking or unverified claims.

#### 4. Mandatory human and visual review
A passing `audit_doc.py --strict` run validates deterministic rules, not argument quality or domain nuance. Always inspect the PNG page previews produced by `render_preview.py` and read the full text from start to finish before sending a document.

#### 5. 100% local execution and privacy
All Python scripts in this repository run locally on your machine with zero outbound network requests. Your `.docx` drafts, logos, and internal figures never leave your computer through these scripts. When interacting with any cloud LLM session, still avoid pasting restricted secrets or confidential records without clearance.

#### 6. Deterministic rules, not blind perfection
Rules deliberately flag terms and patterns heavily overused by LLMs. In documents specifically analyzing AI tools or referencing legitimate names (such as *Claude Lévi-Strauss*), an `ALERTE` may occasionally warrant context-aware judgment rather than mechanical word swapping.

---

<a id="français"></a>

## Français

**Écrire comme un humain, pas comme un modèle de langage.** Encre est un skill éditorial et une suite d’audit `.docx` conçus pour **Claude Code** et **Antigravity**.

Il remplace les réflexes artificiels des assistants (tirets cadratins à répétition, introductions creuses, listes `**Mot-clé** : ...`, sources inventées, bandeaux violets criards, fichiers nommés `document_premier_01.docx`, paramètres de tracking dans les liens et traces techniques dans le XML) par une prose directe, des références vérifiables, une mise en page Word sobre et un contrôle bloquant avant toute remise.

### Exemple Avant / Après

```text
Avant (brouillon généré par défaut) :
Dans un monde en constante évolution — et force est de constater que
c'est fascinant — la transformation digitale n'est pas seulement un outil,
c'est une véritable révolution holistique 🚀. Il convient de souligner que
cet écosystème robuste permet d'optimiser les leviers stratégiques
(source : https://exemple.org/etude?utm_source=chatgpt).

Après (avec Encre) :
Depuis mars 2024, les douze agences régionales reçoivent les factures par
voie numérique. Le délai de traitement est passé de quatre jours à onze
heures sur 18 400 dossiers, principalement grâce à la suppression des
doubles saisies au service d'expédition.
```

### Fonctionnement en six phases

| Phase | Rôle d’Encre | Outil ou référence |
| :--- | :--- | :--- |
| **0. Cadrage avant rédaction** | Pose en une seule fois (3 à 6 questions courtes) les points manquants : auteur, établissement ou entreprise, destinataire, titre exact, logo local, données internes. Aucune balise `[INSÉRER LOGO]` laissée dans le rendu final. | `SKILL.md` |
| **1. Structure fidèle au genre** | Adapte l’architecture au vrai document : page de garde complète et table des matières pour un `memoire`, garde sobre et sommaire pour un `rapport`, en-tête léger sans avant-propos ni sommaire inutile pour un `expose`, bloc direct pour une `note`. | [docx-spec.md](./references/docx-spec.md) |
| **2. Écriture humaine** | Banne les tirets longs (`—`) en incise, les puces mécaniques, les fausses antithèses (`"Ce n'est pas seulement X, mais Y"`), les connecteurs empilés, la sur-justification et le lexique artificiel (FR et EN). | [style-anti-ia.md](./references/style-anti-ia.md) |
| **3. Sources fiables et liens propres** | Privilégie les institutions officielles et revues évaluées par les pairs, interdit les références inventées et purge les paramètres de provenance (`utm_source`, `fbclid`, `ref=`, `=chatgpt`, `=claudecode`, `=gemini`) dans le texte et les hyperliens cachés. | [citations-sources.md](./references/citations-sources.md) |
| **4. Fabrication `.docx` normée** | Produit le document Word avec polices natives (`Times New Roman`, `Calibri`, `Aptos`, `Cambria`, `Garamond`, `Georgia`, `Arial`), palette institutionnelle, vrai champ Table des matières, tableaux propres (`w:tblHeader`, `w:cantSplit`) et XML assaini (`core.xml`, `app.xml`, `custom.xml`, `comments.xml`, `word/_rels/*.rels`). | [build_docx.py](./scripts/build_docx.py), [sanitize_docx.py](./scripts/sanitize_docx.py) |
| **5. Audit bloquant et aperçu visuel** | Contrôle le corps du texte, les cellules de tableaux, les notes, les en-têtes/pieds de page, `word/comments.xml` et les métadonnées XML, puis génère un aperçu PNG page par page via Word ou LibreOffice. | [audit_doc.py](./scripts/audit_doc.py), [render_preview.py](./scripts/render_preview.py) |

### Installation

#### Claude Code

macOS / Linux :

```bash
git clone https://github.com/Sokka04/encre.git ~/.claude/skills/encre
```

Windows (PowerShell) :

```powershell
git clone https://github.com/Sokka04/encre.git "$env:USERPROFILE\.claude\skills\encre"
```

#### Antigravity

macOS / Linux :

```bash
git clone https://github.com/Sokka04/encre.git ~/.gemini/config/skills/encre
```

Windows (PowerShell) :

```powershell
git clone https://github.com/Sokka04/encre.git "$env:USERPROFILE\.gemini\config\skills\encre"
```

Une fois installé, appelle `/encre` ou demande simplement à l’agent de rédiger, réécrire, humaniser ou mettre en forme un mémoire, un rapport, un exposé ou un fichier `.docx`.

### Prérequis et commandes rapides

- Python 3.11+ (`sanitize_docx.py` et `audit_doc.py` fonctionnent avec la seule bibliothèque standard Python)
- `uv` (recommandé ; installe automatiquement `python-docx`, `pillow` et `pymupdf` déclarés dans `build_docx.py` et `render_preview.py`)
- Microsoft Word sous Windows ou LibreOffice (`soffice`) pour convertir le `.docx` en PDF avec actualisation du sommaire puis en images PNG

```bash
# Générer un .docx propre depuis un fichier JSON structuré
uv run scripts/build_docx.py examples/rapport-exemple.json -o "Rapport - Livraisons Tours.docx"

# Assainir métadonnées XML, commentaires et URLs d'un .docx existant
python scripts/sanitize_docx.py "Mon Rapport.docx" --in-place --author "Claire Martin" --title "Audit annuel"

# Lancer l'audit strict (style, tableaux, comments.xml, liens, signatures IA, typographie)
python scripts/audit_doc.py "Rapport - Livraisons Tours.docx" --strict

# Produire les captures PNG page par page pour vérification visuelle
uv run scripts/render_preview.py "Rapport - Livraisons Tours.docx"
```

### Garde-fous et bonnes pratiques

#### 1. Intégrité académique et professionnelle
Encre sert à améliorer la clarté rédactionnelle, la structure logique, la typographie et la propreté des fichiers Word. Il n’est pas destiné au plagiat ni au contournement des règles de transparence d’une université, d’une école ou d’un employeur. L’auteur humain reste entièrement responsable du document remis.

#### 2. Vérification des faits et des citations
Les scripts détectent les liens pollués (`=chatgpt`, `=claudecode`, `=gemini`, `utm_source`), les tics de style et les défauts techniques, mais ne vérifient pas la véracité factuelle d’une affirmation. Contrôle toujours tes chiffres, textes juridiques, citations et références bibliographiques avant diffusion.

#### 3. Portée réelle du nettoyage XML
`sanitize_docx.py` retire les signatures d’outils de génération (`python-docx`, `Pandoc`), nettoie les paramètres de tracking dans les URLs et purge les marqueurs d’assistants dans les propriétés XML et `word/comments.xml` pour livrer un fichier Word propre. Des métadonnées saines ne remplacent jamais un travail de fond rigoureux.

#### 4. Relecture humaine et contrôle visuel obligatoires
Un audit `audit_doc.py --strict` valide des règles déterministes, pas la pertinence d’un raisonnement. Ouvre systématiquement les aperçus PNG générés par `render_preview.py` et relis l’ensemble du document avant tout envoi.

#### 5. Exécution 100 % locale et confidentialité
Tous les scripts Python du dépôt s’exécutent localement sans aucune requête réseau sortante. Tes documents `.docx`, logos et données chiffrées ne transitent par aucun service tiers via ces scripts. Lors des échanges avec un modèle en ligne, évite tout de même d’envoyer des données confidentielles sans autorisation.

#### 6. Règles déterministes et discernement
L’audit signale volontairement les formulations caractéristiques des modèles de langage. Sur un texte traitant justement d’outils numériques ou citant un nom légitime (comme *Claude Lévi-Strauss*), une `ALERTE` appelle un examen humain plutôt qu’un remplacement mécanique.

---

## Licence / License

MIT, voir [LICENSE](./LICENSE).

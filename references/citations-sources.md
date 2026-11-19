# Sources fiables, normes bibliographiques et assainissement d'URL

## 1. Hiérarchie de crédibilité des sources

Avant d'intégrer un chiffre, une date, une référence juridique ou un fait empirique, vérifier l'origine exacte du document :

| Niveau | Sources recommandées | Sources à écarter ou vérifier en amont |
| :--- | :--- | :--- |
| Prioritaire | Institutions publiques et statistiques (INSEE, Banque mondiale, FMI, OMS, OCDE, UNESCO, BCEAO, ministères, Légifrance, EUR-Lex). | Blogs anonymes ou fermes de contenu rédigées pour le référencement. |
| Scientifique | Revues à comité de lecture et archives ouvertes (HAL, Cairn, Persée, PubMed, IEEE, ScienceDirect). | Articles de synthèse générés sans méthodologie ni auteur identifié. |
| Professionnel | Rapports annuels audités, documents techniques de constructeurs, presse économique et générale reconnue. | Forums (Reddit, Quora) cités comme preuve factuelle, pages commerciales promotionnelles. |
| Encyclopédique | Remonter aux sources primaires citées en bas d'article sur Wikipédia. | Ne jamais citer Wikipédia directement comme référence finale d'un mémoire ou rapport. |

Règle absolue : si une référence ne peut pas être vérifiée dans le contexte disponible ou via une recherche web fiable, ne jamais reconstituer de mémoire un numéro de page, un titre d'article, un DOI ou une URL. Signaler le manque à l'utilisateur ou demander ses documents internes.

## 2. Modèles de présentation bibliographique

### Norme APA 7 (auteur-date)

- **Ouvrage** : Nom, P. (Année). *Titre de l'ouvrage en italique* (éd.). Éditeur.
- **Article scientifique** : Nom, P., & Nom2, P. (Année). Titre de l'article. *Nom de la revue*, *Volume*(Numéro), p. début-p. fin. https://doi.org/...
- **Rapport institutionnel** : Organisme. (Année). *Titre du rapport*. URL canonique propre.
- **Texte juridique (France)** : Loi n° YYYY-XXXX du JJ mois AAAA relative à... *Journal officiel de la République française*. https://www.legifrance.gouv.fr/...

### Norme ISO 690 (française)

- **Ouvrage** : NOM, Prénom. *Titre de l'ouvrage*. Ville d'édition : Éditeur, année. Pagination.
- **Article en ligne** : NOM, Prénom. Titre de l'article. *Titre de la revue*, année, vol. X, n° Y, p. Z-W. Disponible sur : URL canonique propre (consulté le JJ/MM/AAAA).

## 3. Assainissement obligatoire des URLs

Toute URL insérée dans le texte, en note de bas de page, en bibliographie ou en lien hypertexte Word caché (`word/_rels/document.xml.rels`) doit être réduite à sa forme canonique.

Paramètres à supprimer systématiquement :

1. Marqueurs de campagne et réseaux sociaux : `utm_source`, `utm_medium`, `utm_campaign`, `utm_term`, `utm_content`, `utm_id`, `fbclid`, `gclid`, `mc_cid`, `mc_eid`, `igshid`.
2. Signatures d'assistants et moteurs : `chatgpt`, `openai`, `claude`, `claudecode`, `anthropic`, `gemini`, `perplexity`, `copilot`, `bing`.
3. Paramètres de parrainage ou provenance : `ref=`.

Le script `scripts/sanitize_docx.py` et l'étape post-génération de `scripts/build_docx.py` effectuent cette purge automatiquement dans le texte XML et dans les relations externes du `.docx`.

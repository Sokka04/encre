"""Tests unitaires et d'intégration pour les scripts du skill Encre."""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

ROOT_DIR = Path(__file__).resolve().parent.parent
SCRIPTS_DIR = ROOT_DIR / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

import audit_doc
import build_docx
import sanitize_docx

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"


class TestEncreSuite(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory(prefix="encre_test_")
        self.tmp_dir = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_clean_url_preserves_legit_params(self) -> None:
        raw = "https://www.insee.fr/fr/statistiques/42?id=99&utm_source=chatgpt.com&fbclid=xyz&ref=claudecode"
        cleaned, changed = sanitize_docx.clean_url(raw)
        self.assertTrue(changed)
        self.assertEqual(cleaned, "https://www.insee.fr/fr/statistiques/42?id=99")

        clean_already, changed2 = sanitize_docx.clean_url(
            "https://www.legifrance.gouv.fr/loda/id/JORFTEXT0000"
        )
        self.assertFalse(changed2)
        self.assertEqual(
            clean_already, "https://www.legifrance.gouv.fr/loda/id/JORFTEXT0000"
        )

    def test_build_all_four_genres(self) -> None:
        example_spec = json.loads(
            (ROOT_DIR / "examples" / "rapport-exemple.json").read_text(encoding="utf-8")
        )

        # 1. Rapport (exemple complet)
        rapport_path = self.tmp_dir / "Rapport - Livraisons Tours.docx"
        build_docx.build_docx(example_spec, output_path=rapport_path)
        self.assertTrue(rapport_path.exists())

        report = audit_doc.audit_document(rapport_path, lang="fr")
        self.assertEqual(report["metrics"]["errors"], 0, report["findings"])
        self.assertEqual(report["metrics"]["alerts"], 0, report["findings"])
        self.assertTrue(report["passed_strict"])

        with zipfile.ZipFile(rapport_path, "r") as zf:
            core_xml = zf.read("docProps/core.xml").decode("utf-8")
            app_xml = zf.read("docProps/app.xml").decode("utf-8")
            settings_xml = zf.read("word/settings.xml").decode("utf-8")
            doc_xml = zf.read("word/document.xml").decode("utf-8")

        self.assertNotIn("python-docx", core_xml.lower())
        self.assertIn("Camille Bernard", core_xml)
        self.assertIn("Microsoft Office Word", app_xml)
        self.assertIn("updateFields", settings_xml)
        self.assertIn("w:tblHeader", doc_xml)
        self.assertIn("w:cantSplit", doc_xml)
        self.assertIn("w:titlePg", doc_xml)
        self.assertIn("«\u00a0", doc_xml)

        # 2. Mémoire avec pages liminaires en chiffres romains
        memoire_spec = {
            "genre": "memoire",
            "title": "Gestion des flux logistiques urbains",
            "organization": "Université de Tours",
            "author": "Camille Bernard",
            "supervisor": "Pr. Claire Moreau",
            "program": "Master Transport et Mobilité",
            "academic_year": "2026-2027",
            "preliminaries": [
                {
                    "title": "Résumé",
                    "blocks": [
                        {
                            "type": "paragraph",
                            "text": "Ce mémoire étudie l'organisation des tournées matinales dans trois agglomérations moyennes.",
                        }
                    ],
                }
            ],
            "sections": [
                {
                    "level": 1,
                    "title": "Introduction",
                    "blocks": [
                        {
                            "type": "paragraph",
                            "text": "Les livraisons du matin concentrent 68 % des stationnements en double file entre 8 h et 10 h.",
                        }
                    ],
                }
            ],
        }
        memoire_path = self.tmp_dir / "Memoire - Flux logistiques - Bernard.docx"
        build_docx.build_docx(memoire_spec, output_path=memoire_path)
        with zipfile.ZipFile(memoire_path, "r") as zf:
            mdoc = zf.read("word/document.xml").decode("utf-8")
        self.assertIn('w:fmt="lowerRoman"', mdoc)
        self.assertIn('w:fmt="decimal"', mdoc)
        self.assertEqual(
            audit_doc.audit_document(memoire_path, lang="fr")["metrics"]["errors"], 0
        )

        # 3. Exposé sans page de garde ni sommaire
        expose_spec = {
            "genre": "expose",
            "title": "Fonctionnement des zones à faibles émissions",
            "organization": "Licence Aménagement",
            "author": "Camille Bernard",
            "date": "18 novembre 2026",
            "sections": [
                {
                    "level": 1,
                    "title": "1. Cadre réglementaire",
                    "blocks": [
                        {
                            "type": "paragraph",
                            "text": "La loi fixe des seuils progressifs pour les véhicules utilitaires légers.",
                        }
                    ],
                }
            ],
        }
        expose_path = self.tmp_dir / "Expose - Zones faibles emissions.docx"
        build_docx.build_docx(expose_spec, output_path=expose_path)
        with zipfile.ZipFile(expose_path, "r") as zf:
            edoc = zf.read("word/document.xml").decode("utf-8")
        self.assertNotIn("TOC \\o", edoc)
        self.assertEqual(
            audit_doc.audit_document(expose_path, lang="fr")["metrics"]["errors"], 0
        )

        # 4. Note administrative sans page de garde ni pagination
        note_spec = {
            "genre": "note",
            "title": "Modification des horaires d'ouverture du quai B",
            "organization": "Service Exploitation",
            "author": "Camille Bernard",
            "recipient": "Chefs d'équipe",
            "date": "18 novembre 2026",
            "reference": "NOTE-2026-08",
            "sections": [
                {
                    "level": 1,
                    "title": "Nouvelles plages horaires",
                    "blocks": [
                        {
                            "type": "paragraph",
                            "text": "À compter du 1er décembre 2026, le quai B ouvre dès 5 h 45 du lundi au vendredi.",
                        }
                    ],
                }
            ],
        }
        note_path = self.tmp_dir / "Note - Horaires quai B.docx"
        build_docx.build_docx(note_spec, output_path=note_path)
        self.assertEqual(
            audit_doc.audit_document(note_path, lang="fr")["metrics"]["errors"], 0
        )

    def test_build_rejects_exotic_font_ai_color_and_placeholder(self) -> None:
        base = {"genre": "note", "title": "Test validation", "sections": []}

        bad_font = dict(base, font="Comic Sans MS")
        with self.assertRaises(ValueError):
            build_docx.build_docx(bad_font, output_path=self.tmp_dir / "Bad Font.docx")

        bad_color = dict(base, palette="#6366F1")
        with self.assertRaises(ValueError):
            build_docx.build_docx(
                bad_color, output_path=self.tmp_dir / "Bad Color.docx"
            )

        bad_placeholder = {
            "genre": "note",
            "title": "Test",
            "sections": [
                {
                    "level": 1,
                    "title": "Axe",
                    "blocks": [{"type": "paragraph", "text": "[INSÉRER LOGO]"}],
                }
            ],
        }
        with self.assertRaises(ValueError):
            build_docx.build_docx(
                bad_placeholder, output_path=self.tmp_dir / "Bad Placeholder.docx"
            )

    def test_sanitize_poisoned_docx(self) -> None:
        poisoned = self.tmp_dir / "Document empoisonne.docx"
        doc = Document()
        doc.core_properties.author = "python-docx generator"
        doc.core_properties.last_modified_by = "Claude Assistant"
        doc.core_properties.comments = "Generated by python-docx and OpenAI"

        p = doc.add_paragraph(
            "Consulter https://insee.fr/fr/statistiques/123?ref=claudecode pour les chiffres."
        )
        r_id = p.part.relate_to(
            "https://example.org/rapport?utm_source=chatgpt.com&fbclid=abc&page=2",
            RELATIONSHIP_TYPE.HYPERLINK,
            is_external=True,
        )
        hl = OxmlElement("w:hyperlink")
        hl.set(qn("r:id"), r_id)
        run = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = "Lien source"
        run.append(t)
        hl.append(run)
        p._p.append(hl)
        doc.save(str(poisoned))

        # Injecte aussi Pandoc dans docProps/app.xml pour simuler un fichier contaminé
        patched = self.tmp_dir / "Document empoisonne Pandoc.docx"
        with (
            zipfile.ZipFile(poisoned, "r") as zin,
            zipfile.ZipFile(patched, "w") as zout,
        ):
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "docProps/app.xml":
                    data = data.replace(b"python-docx", b"Pandoc")
                    if b"Pandoc" not in data:
                        data = data.replace(b"</Application>", b" Pandoc</Application>")
                zout.writestr(item, data)

        summary = sanitize_docx.sanitize_docx(
            patched,
            author="Claire Martin",
            title="Rapport assaini",
            in_place=True,
        )
        self.assertGreaterEqual(summary["total_urls_cleaned"], 2)

        # Vérifie l'absence totale d'empreintes et la réouverture valide
        reopened = Document(str(patched))
        self.assertEqual(reopened.core_properties.author, "Claire Martin")
        self.assertEqual(reopened.core_properties.title, "Rapport assaini")

        with zipfile.ZipFile(patched, "r") as zf:
            for name in zf.namelist():
                if name.endswith((".xml", ".rels")):
                    ET.fromstring(zf.read(name))
            core = zf.read("docProps/core.xml").decode("utf-8")
            app = zf.read("docProps/app.xml").decode("utf-8")
            rels = zf.read("word/_rels/document.xml.rels").decode("utf-8")
            dxml = zf.read("word/document.xml").decode("utf-8")

        for leak in (
            "python-docx",
            "Pandoc",
            "chatgpt",
            "claudecode",
            "utm_source",
            "fbclid",
            "ref=",
        ):
            self.assertNotIn(leak.lower(), (core + app + rels + dxml).lower())
        self.assertIn("https://example.org/rapport?page=2", rels)
        self.assertIn("https://insee.fr/fr/statistiques/123", dxml)

    def test_audit_catches_ai_and_technical_defects(self) -> None:
        bad_md = self.tmp_dir / "brouillon.md"
        bad_md.write_text(
            "# Transformation digitale : un levier stratégique majeur\n\n"
            "Dans un monde en constante évolution — et c'est fascinant — il convient de noter que ce n'est pas seulement un outil, mais une révolution holistique 🚀.\n\n"
            "De plus, cette approche permet d'optimiser la rentabilité.\n\n"
            "Par ailleurs, elle reste crucial et crucial pour le paysage numérique.\n\n"
            "- **Rapidité** : traitement accéléré des dossiers.\n\n"
            "Voir [VOTRE NOM] sur https://example.org?utm_source=chatgpt.\n\n"
            "En conclusion, l'étude ouvre la voie à un avenir prometteur.\n",
            encoding="utf-8",
        )
        rep_md = audit_doc.audit_document(bad_md, lang="fr")
        codes = {f["code"] for f in rep_md["findings"]}
        for expected_code in (
            "em_dash",
            "emoji",
            "bold_lead_bullet",
            "dramatic_colon_heading",
            "false_antithesis",
            "meta_commentary",
            "generic_framing",
            "consecutive_connectors",
            "ai_slop_lexicon",
            "repeated_ai_buzzword",
            "dirty_url_text",
            "placeholder_text",
        ):
            self.assertIn(
                expected_code, codes, f"Code attendu manquant : {expected_code}"
            )
        self.assertFalse(rep_md["passed_default"])

        # Test négatif .docx : nom robotique, markdown brut, couleur #6366F1, puce gras-deux-points
        bad_docx = self.tmp_dir / "rapport_v1_final.docx"
        d = Document()
        d.add_paragraph("**Texte brut** non converti en gras Word.")
        p_bullet = d.add_paragraph(style="List Bullet")
        r1 = p_bullet.add_run("Critère :")
        r1.bold = True
        p_bullet.add_run(" explication mécanique.")
        p_col = d.add_paragraph()
        p_col._p.get_or_add_pPr().append(
            parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:fill="6366F1"/>')
        )
        p_col.add_run("Bloc violet IA.")
        d.save(str(bad_docx))

        rep_docx = audit_doc.audit_document(bad_docx, lang="fr")
        dcodes = {f["code"] for f in rep_docx["findings"]}
        for expected_code in (
            "robotic_filename",
            "raw_markdown",
            "bold_lead_bullet",
            "forbidden_ai_color",
            "xml_metadata_leak",
        ):
            self.assertIn(
                expected_code, dcodes, f"Code .docx attendu manquant : {expected_code}"
            )

    def test_audit_no_false_positives_on_legit_names_and_times(self) -> None:
        clean_md = self.tmp_dir / "note-propre.md"
        clean_md.write_text(
            "# Tableau 1\u00a0: Répartition des effectifs\n\n"
            "Les travaux de Claude Lévi-Strauss restent cités dans les programmes universitaires. "
            "La réunion débute à 14:30 dans la salle B, puis les équipes présentent les comptes trimestriels.\n",
            encoding="utf-8",
        )
        rep = audit_doc.audit_document(clean_md, lang="fr")
        self.assertEqual(rep["metrics"]["errors"], 0, rep["findings"])
        self.assertEqual(rep["metrics"]["alerts"], 0, rep["findings"])

    def test_audit_catches_defects_inside_docx_tables(self) -> None:
        bad_tbl_docx = self.tmp_dir / "Tableau defectueux.docx"
        d = Document()
        tbl = d.add_table(rows=2, cols=3)
        tbl.cell(0, 0).text = "Critère — 1"
        tbl.cell(0, 1).text = "Signal 🚀"
        tbl.cell(0, 2).text = "[VOTRE NOM]"
        tbl.cell(1, 0).text = "**gras brut**"
        tbl.cell(1, 1).text = "https://insee.fr?utm_source=chatgpt.com"
        tbl.cell(1, 2).text = "source=claudecode"
        d.save(str(bad_tbl_docx))

        rep = audit_doc.audit_document(bad_tbl_docx, lang="fr")
        table_codes = {
            f["code"]
            for f in rep["findings"]
            if f["location"].startswith("Tableau 1,")
        }
        for expected_code in (
            "em_dash",
            "emoji",
            "placeholder_text",
            "raw_markdown",
            "dirty_url_text",
            "ai_signature_text",
        ):
            self.assertIn(
                expected_code,
                table_codes,
                f"Code attendu manquant dans le tableau : {expected_code}",
            )

    def test_audit_ai_signatures_outside_urls_and_no_false_positives(self) -> None:
        bad_sig_md = self.tmp_dir / "signatures-ia.md"
        bad_sig_md.write_text(
            "Source : ChatGPT\n\n"
            "Référence issue de insee.fr?utm_source=gemini.\n\n"
            "Paramètre =claudecode.\n\n"
            "Généré par Claude Code.\n",
            encoding="utf-8",
        )
        rep_bad = audit_doc.audit_document(bad_sig_md, lang="fr")
        sig_findings = [
            f for f in rep_bad["findings"] if f["code"] == "ai_signature_text"
        ]
        self.assertEqual(len(sig_findings), 4, rep_bad["findings"])

        clean_sig_md = self.tmp_dir / "etude-propre.md"
        clean_sig_md.write_text(
            "# Analyse des usages numériques\n\n"
            "Les travaux de Claude Lévi-Strauss restent étudiés en anthropologie.\n\n"
            "L’enquête analyse l’usage de ChatGPT par les étudiants en licence dans trois universités.\n",
            encoding="utf-8",
        )
        rep_clean = audit_doc.audit_document(clean_sig_md, lang="fr")
        self.assertEqual(rep_clean["metrics"]["errors"], 0, rep_clean["findings"])
        self.assertEqual(rep_clean["metrics"]["alerts"], 0, rep_clean["findings"])

    def test_sanitize_and_audit_word_comments_xml(self) -> None:
        base_docx = self.tmp_dir / "Base note.docx"
        d = Document()
        d.add_paragraph("Premier paragraphe propre.")
        d.save(str(base_docx))

        commented = self.tmp_dir / "Note commentee.docx"
        comments_payload = (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            b'<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            b'<w:comment w:id="0" w:author="Claude Assistant" w:initials="CA">'
            b"<w:p><w:r><w:t>"
            b"Voir https://www.insee.fr/fr/statistiques/12?utm_source=gemini&amp;page=1"
            b"</w:t></w:r></w:p>"
            b"</w:comment>"
            b"</w:comments>"
        )

        with (
            zipfile.ZipFile(base_docx, "r") as zin,
            zipfile.ZipFile(commented, "w") as zout,
        ):
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "[Content_Types].xml":
                    override = (
                        b'<Override PartName="/word/comments.xml" '
                        b'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'
                    )
                    data = data.replace(b"</Types>", override + b"</Types>")
                zout.writestr(item, data)
            zout.writestr("word/comments.xml", comments_payload)

        rep_before = audit_doc.audit_document(commented, lang="fr")
        comment_codes_before = {
            f["code"]
            for f in rep_before["findings"]
            if f["location"] == "word/comments.xml"
        }
        self.assertIn("xml_metadata_leak", comment_codes_before, rep_before["findings"])
        self.assertIn("dirty_url_extra", comment_codes_before, rep_before["findings"])

        summary = sanitize_docx.sanitize_docx(
            commented,
            author="Claire Martin",
            title="Note revue",
            in_place=True,
        )
        self.assertIn("comments:author", summary["metadata_changes"])
        self.assertGreaterEqual(summary["urls_cleaned_xml"], 1)

        with zipfile.ZipFile(commented, "r") as zf:
            for name in zf.namelist():
                if name.endswith((".xml", ".rels")):
                    ET.fromstring(zf.read(name))
            comments_xml = zf.read("word/comments.xml").decode("utf-8")

        self.assertNotIn("Claude Assistant", comments_xml)
        self.assertNotIn("utm_source", comments_xml)
        self.assertIn("Claire Martin", comments_xml)
        self.assertIn('w:initials="CM"', comments_xml)
        self.assertIn("https://www.insee.fr/fr/statistiques/12?page=1", comments_xml)

        rep_after = audit_doc.audit_document(commented, lang="fr")
        comment_findings_after = [
            f for f in rep_after["findings"] if f["location"] == "word/comments.xml"
        ]
        self.assertEqual(comment_findings_after, [])
        self.assertEqual(rep_after["metrics"]["errors"], 0, rep_after["findings"])


if __name__ == "__main__":
    unittest.main()

# /// script
# requires-python = ">=3.11"
# dependencies = [
#   "python-docx>=1.1.2",
#   "pillow>=10.0.0",
# ]
# ///
"""Génère un fichier Word (.docx) sobre, normé et sans marqueurs IA depuis un JSON structuré."""

from __future__ import annotations

import argparse
import json
import re
import struct
import sys
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, RGBColor

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from sanitize_docx import (
    clean_url,
    sanitize_docx,
    sanitize_filename,
    split_trailing_punctuation,
)

try:
    from PIL import Image  # type: ignore[import-not-found]
except ImportError:
    Image = None

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

NATIVE_FONTS = {
    "Times New Roman",
    "Garamond",
    "Georgia",
    "Calibri",
    "Aptos",
    "Cambria",
    "Arial",
}

PALETTES = {
    "monochrome": {
        "accent": "1A1A1A",
        "text": "1A1A1A",
        "muted": "4B5563",
        "table_header": "F3F4F6",
        "border": "D1D5DB",
        "border_strong": "9CA3AF",
    },
    "bleu-nuit": {
        "accent": "1B365D",
        "text": "1A1A1A",
        "muted": "4B5563",
        "table_header": "F3F4F6",
        "border": "D1D5DB",
        "border_strong": "9CA3AF",
    },
    "ardoise": {
        "accent": "2C3E50",
        "text": "1A1A1A",
        "muted": "4B5563",
        "table_header": "F3F4F6",
        "border": "D1D5DB",
        "border_strong": "9CA3AF",
    },
    "vert-foret": {
        "accent": "1E4D3B",
        "text": "1A1A1A",
        "muted": "4B5563",
        "table_header": "F3F4F6",
        "border": "D1D5DB",
        "border_strong": "9CA3AF",
    },
    "bordeaux": {
        "accent": "6B2737",
        "text": "1A1A1A",
        "muted": "4B5563",
        "table_header": "F3F4F6",
        "border": "D1D5DB",
        "border_strong": "9CA3AF",
    },
}

FORBIDDEN_AI_COLORS = {
    "6366F1",
    "7C3AED",
    "A855F7",
    "8B5CF6",
    "06B6D4",
    "4F46E5",
    "9333EA",
    "3B82F6",
}

GENRE_DEFAULTS: dict[str, dict[str, Any]] = {
    "memoire": {
        "font": "Times New Roman",
        "body_size": 12.0,
        "line_spacing": 1.5,
        "margins_cm": {"top": 2.5, "bottom": 2.5, "left": 3.0, "right": 2.5},
        "cover_page": True,
        "toc": True,
        "toc_title": "Table des matières",
        "header": True,
        "page_numbers": True,
        "palette": "monochrome",
        "prefix": "Memoire",
    },
    "rapport": {
        "font": "Calibri",
        "body_size": 11.0,
        "line_spacing": 1.15,
        "margins_cm": {"top": 2.5, "bottom": 2.5, "left": 2.5, "right": 2.5},
        "cover_page": True,
        "toc": True,
        "toc_title": "Sommaire",
        "header": True,
        "page_numbers": True,
        "palette": "bleu-nuit",
        "prefix": "Rapport",
    },
    "expose": {
        "font": "Calibri",
        "body_size": 11.0,
        "line_spacing": 1.15,
        "margins_cm": {"top": 2.5, "bottom": 2.5, "left": 2.5, "right": 2.5},
        "cover_page": False,
        "toc": False,
        "toc_title": "Sommaire",
        "header": True,
        "page_numbers": True,
        "palette": "bleu-nuit",
        "prefix": "Expose",
    },
    "note": {
        "font": "Calibri",
        "body_size": 11.0,
        "line_spacing": 1.15,
        "margins_cm": {"top": 2.5, "bottom": 2.5, "left": 2.5, "right": 2.5},
        "cover_page": False,
        "toc": False,
        "toc_title": "Sommaire",
        "header": False,
        "page_numbers": False,
        "palette": "bleu-nuit",
        "prefix": "Note",
    },
}

PLACEHOLDER_PATTERNS = (
    "[insérer",
    "[inserer",
    "[votre",
    "[nom",
    "[date",
    "[logo",
    "lorem ipsum",
)

URL_IN_TEXT_RE = re.compile(r"https?://[^\s<>\"']+")


def _hex_to_rgb(hex_str: str) -> RGBColor:
    cleaned = hex_str.lstrip("#").upper()
    return RGBColor(int(cleaned[0:2], 16), int(cleaned[2:4], 16), int(cleaned[4:6], 16))


def apply_french_typography(text: str, lang: str = "fr") -> str:
    """Applique les règles typographiques françaises tout en préservant les URLs."""
    if not text:
        return text

    urls: list[str] = []

    def _stash_url(match: re.Match[str]) -> str:
        core, suffix = split_trailing_punctuation(match.group(0))
        cleaned, _ = clean_url(core)
        token = f"\x00URL{len(urls)}\x00"
        urls.append(cleaned)
        return f"{token}{suffix}"

    result = URL_IN_TEXT_RE.sub(_stash_url, text)

    if lang.lower().startswith("fr"):
        result = result.replace("'", "\u2019")
        result = re.sub(
            r'"([^"\n]+)"', lambda m: f"«\u00a0{m.group(1).strip()}\u00a0»", result
        )
        result = re.sub(r"«[ \t\u00A0\u202F]*", "«\u00a0", result)
        result = re.sub(r"[ \t\u00A0\u202F]*»", "\u00a0»", result)
        result = re.sub(
            r"([A-Za-zÀ-ÖØ-öø-ÿ0-9»)])[ \t\u00A0\u202F]*([;!?])(?=\s|$)",
            "\\1\u00a0\\2",
            result,
        )
        result = re.sub(
            r"([A-Za-zÀ-ÖØ-öø-ÿ0-9»)])[ \t\u00A0\u202F]+:(?=\s|$)",
            "\\1\u00a0:",
            result,
        )
        result = re.sub(
            r"([A-Za-zÀ-ÖØ-öø-ÿ»)]):(?=\s)",
            "\\1\u00a0:",
            result,
        )
        result = re.sub(r"(\d)[ \t\u00A0\u202F]*([%€])", "\\1\u00a0\\2", result)

    for idx, url in enumerate(urls):
        result = result.replace(f"\x00URL{idx}\x00", url)
    return result


def _scan_placeholders(value: Any, path: str = "root") -> list[str]:
    hits: list[str] = []
    if isinstance(value, str):
        low = value.lower()
        for pat in PLACEHOLDER_PATTERNS:
            if pat in low:
                hits.append(f"{path}: contient « {pat} »")
    elif isinstance(value, list):
        for idx, item in enumerate(value):
            hits.extend(_scan_placeholders(item, f"{path}[{idx}]"))
    elif isinstance(value, dict):
        for key, item in value.items():
            hits.extend(_scan_placeholders(item, f"{path}.{key}"))
    return hits


def _resolve_palette(
    palette_spec: str | None, default_name: str, allow_custom: bool
) -> dict[str, str]:
    if not palette_spec:
        return dict(PALETTES[default_name])

    key = palette_spec.strip().lower()
    if key in PALETTES:
        return dict(PALETTES[key])

    match = re.fullmatch(r"#?([0-9a-fA-F]{6})", palette_spec.strip())
    if not match:
        raise ValueError(
            f"Palette inconnue : {palette_spec!r}. "
            f"Choix valides : {', '.join(PALETTES)} ou code hexadécimal #RRGGBB."
        )

    hex_code = match.group(1).upper()
    if hex_code in FORBIDDEN_AI_COLORS and not allow_custom:
        raise ValueError(
            f"Couleur #{hex_code} refusée (palette IA générique). "
            "Utilisez une teinte institutionnelle sobre ou --allow-custom-color."
        )
    palette = dict(PALETTES["monochrome"])
    palette["accent"] = hex_code
    return palette


def _set_run_fonts_and_lang(r_pr: OxmlElement, font_name: str, lang_tag: str) -> None:
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, r_fonts)
    for theme_attr in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
        r_fonts.attrib.pop(qn(theme_attr), None)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        r_fonts.set(qn(attr), font_name)

    lang_elem = r_pr.find(qn("w:lang"))
    if lang_elem is None:
        lang_elem = OxmlElement("w:lang")
        r_pr.append(lang_elem)
    lang_elem.set(qn("w:val"), lang_tag)


def _set_paragraph_bottom_border(
    paragraph: Any, color_hex: str, size_eighths: int = 6, space: int = 6
) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = p_pr.find(qn("w:pBdr"))
    if p_bdr is None:
        p_bdr = OxmlElement("w:pBdr")
        p_pr.append(p_bdr)
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size_eighths))
    bottom.set(qn("w:space"), str(space))
    bottom.set(qn("w:color"), color_hex)
    p_bdr.append(bottom)


def _set_paragraph_left_border(
    paragraph: Any, color_hex: str, size_eighths: int = 16, space: int = 10
) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = p_pr.find(qn("w:pBdr"))
    if p_bdr is None:
        p_bdr = OxmlElement("w:pBdr")
        p_pr.append(p_bdr)
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), str(size_eighths))
    left.set(qn("w:space"), str(space))
    left.set(qn("w:color"), color_hex)
    p_bdr.append(left)


def _configure_styles(
    doc: Document,
    *,
    font_name: str,
    body_size: float,
    line_spacing: float,
    palette: dict[str, str],
    genre: str,
    lang: str,
) -> None:
    lang_tag = "fr-FR" if lang.lower().startswith("fr") else "en-US"
    text_rgb = _hex_to_rgb(palette["text"])
    accent_rgb = _hex_to_rgb(palette["accent"])
    muted_rgb = _hex_to_rgb(palette["muted"])

    normal = doc.styles["Normal"]
    normal.font.name = font_name
    normal.font.size = Pt(body_size)
    normal.font.color.rgb = text_rgb
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    normal.paragraph_format.line_spacing = line_spacing
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.widow_control = True
    _set_run_fonts_and_lang(normal.element.get_or_add_rPr(), font_name, lang_tag)

    h1_size = 16.0 if genre in {"memoire", "rapport"} else 15.0
    heading_specs = [
        ("Heading 1", h1_size, True, False, 16, 6, accent_rgb),
        ("Heading 2", 13.0, True, False, 12, 4, accent_rgb),
        ("Heading 3", 11.5, True, False, 9, 3, text_rgb),
    ]
    for (
        style_name,
        pt_size,
        bold,
        italic,
        before_pt,
        after_pt,
        color_rgb,
    ) in heading_specs:
        style = doc.styles[style_name]
        style.font.name = font_name
        style.font.size = Pt(pt_size)
        style.font.bold = bold
        style.font.italic = italic
        style.font.color.rgb = color_rgb
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        style.paragraph_format.line_spacing = 1.15
        style.paragraph_format.space_before = Pt(before_pt)
        style.paragraph_format.space_after = Pt(after_pt)
        style.paragraph_format.keep_with_next = True
        style.paragraph_format.keep_together = True
        _set_run_fonts_and_lang(style.element.get_or_add_rPr(), font_name, lang_tag)

    if genre == "rapport":
        p_pr = doc.styles["Heading 1"].element.get_or_add_pPr()
        p_bdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "4")
        bottom.set(qn("w:space"), "3")
        bottom.set(qn("w:color"), palette["border"])
        p_bdr.append(bottom)
        p_pr.append(p_bdr)

    for list_style_name in ("List Bullet", "List Number"):
        if list_style_name in doc.styles:
            lstyle = doc.styles[list_style_name]
            lstyle.font.name = font_name
            lstyle.font.size = Pt(body_size)
            lstyle.font.color.rgb = text_rgb
            lstyle.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            lstyle.paragraph_format.line_spacing = line_spacing
            lstyle.paragraph_format.space_after = Pt(3)
            lstyle.paragraph_format.widow_control = True
            _set_run_fonts_and_lang(
                lstyle.element.get_or_add_rPr(), font_name, lang_tag
            )

    if "Caption" in doc.styles:
        cstyle = doc.styles["Caption"]
        cstyle.font.name = font_name
        cstyle.font.size = Pt(9.5)
        cstyle.font.bold = True
        cstyle.font.italic = False
        cstyle.font.color.rgb = muted_rgb
        cstyle.paragraph_format.keep_with_next = True
        _set_run_fonts_and_lang(cstyle.element.get_or_add_rPr(), font_name, lang_tag)


def _enable_update_fields_on_open(doc: Document) -> None:
    settings_elem = doc.settings.element
    existing = settings_elem.find(qn("w:updateFields"))
    if existing is None:
        uf = OxmlElement("w:updateFields")
        uf.set(qn("w:val"), "true")
        settings_elem.append(uf)
    else:
        existing.set(qn("w:val"), "true")


def _set_section_margins(section: Any, margins_cm: dict[str, float]) -> None:
    section.top_margin = Cm(float(margins_cm.get("top", 2.5)))
    section.bottom_margin = Cm(float(margins_cm.get("bottom", 2.5)))
    section.left_margin = Cm(float(margins_cm.get("left", 2.5)))
    section.right_margin = Cm(float(margins_cm.get("right", 2.5)))
    section.header_distance = Cm(1.25)
    section.footer_distance = Cm(1.25)


def _set_section_page_number_format(
    section: Any, fmt: str = "decimal", start: int | None = None
) -> None:
    sect_pr = section._sectPr
    pg_num = sect_pr.find(qn("w:pgNumType"))
    if pg_num is None:
        pg_num = OxmlElement("w:pgNumType")
        sect_pr.append(pg_num)
    pg_num.set(qn("w:fmt"), fmt)
    if start is not None:
        pg_num.set(qn("w:start"), str(start))


def _append_field(
    paragraph: Any, instr: str, font_name: str, size_pt: float, color_hex: str
) -> None:
    def _styled_rpr() -> OxmlElement:
        r_pr = OxmlElement("w:rPr")
        r_fonts = OxmlElement("w:rFonts")
        for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            r_fonts.set(qn(attr), font_name)
        r_pr.append(r_fonts)
        col = OxmlElement("w:color")
        col.set(qn("w:val"), color_hex)
        r_pr.append(col)
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), str(int(size_pt * 2)))
        r_pr.append(sz)
        sz_cs = OxmlElement("w:szCs")
        sz_cs.set(qn("w:val"), str(int(size_pt * 2)))
        r_pr.append(sz_cs)
        return r_pr

    run_begin = OxmlElement("w:r")
    run_begin.append(_styled_rpr())
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    run_begin.append(fld_begin)

    run_instr = OxmlElement("w:r")
    run_instr.append(_styled_rpr())
    instr_el = OxmlElement("w:instrText")
    instr_el.set(qn("xml:space"), "preserve")
    instr_el.text = f" {instr} "
    run_instr.append(instr_el)

    run_sep = OxmlElement("w:r")
    run_sep.append(_styled_rpr())
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    run_sep.append(fld_sep)

    run_val = OxmlElement("w:r")
    run_val.append(_styled_rpr())
    txt = OxmlElement("w:t")
    txt.text = "1"
    run_val.append(txt)

    run_end = OxmlElement("w:r")
    run_end.append(_styled_rpr())
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run_end.append(fld_end)

    for r in (run_begin, run_instr, run_sep, run_val, run_end):
        paragraph._p.append(r)


def _configure_headers_and_footers(
    section: Any,
    *,
    show_header: bool,
    show_page_numbers: bool,
    different_first_page: bool,
    short_title: str,
    font_name: str,
    palette: dict[str, str],
) -> None:
    section.different_first_page_header_footer = different_first_page
    muted_rgb = _hex_to_rgb(palette["muted"])

    if different_first_page:
        for container in (section.first_page_header, section.first_page_footer):
            container.is_linked_to_previous = False
            for p in container.paragraphs:
                p.text = ""

    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    hp.text = ""
    if show_header and short_title:
        hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
        hp.paragraph_format.space_after = Pt(4)
        hrun = hp.add_run(short_title)
        hrun.font.name = font_name
        hrun.font.size = Pt(9)
        hrun.font.color.rgb = muted_rgb
        _set_paragraph_bottom_border(hp, palette["border"], size_eighths=4, space=4)

    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.text = ""
    if show_page_numbers:
        fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        fp.paragraph_format.space_before = Pt(4)
        _append_field(fp, "PAGE", font_name, 9.5, palette["muted"])


def _get_image_dimensions_px(image_path: Path) -> tuple[int, int]:
    if Image is not None:
        with Image.open(image_path) as img:
            return img.size

    raw = image_path.read_bytes()
    if raw[:8] == b"\x89PNG\r\n\x1a\n" and len(raw) >= 24:
        w, h = struct.unpack(">II", raw[16:24])
        return int(w), int(h)
    if raw[:2] == b"\xff\xd8":
        idx = 2
        while idx + 9 < len(raw):
            if raw[idx] != 0xFF:
                break
            marker = raw[idx + 1]
            if marker in (0xC0, 0xC1, 0xC2, 0xC3):
                h, w = struct.unpack(">HH", raw[idx + 5 : idx + 9])
                return int(w), int(h)
            seg_len = struct.unpack(">H", raw[idx + 2 : idx + 4])[0]
            idx += 2 + seg_len
    return 1000, 500


def _resolve_asset_path(raw_path: str, base_dir: Path) -> Path:
    candidate = Path(raw_path)
    if candidate.is_absolute() and candidate.exists():
        return candidate
    rel = (base_dir / candidate).resolve()
    if rel.exists():
        return rel
    cwd_rel = (Path.cwd() / candidate).resolve()
    if cwd_rel.exists():
        return cwd_rel
    raise FileNotFoundError(f"Fichier média introuvable : {raw_path}")


def _add_hyperlink(
    paragraph: Any,
    text: str,
    url: str,
    *,
    font_name: str,
    size_pt: float,
    color_hex: str,
    bold: bool = False,
    italic: bool = False,
) -> None:
    cleaned_url, _ = clean_url(url)
    r_id = paragraph.part.relate_to(
        cleaned_url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True
    )

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    new_run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    r_fonts = OxmlElement("w:rFonts")
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        r_fonts.set(qn(attr), font_name)
    r_pr.append(r_fonts)

    if bold:
        r_pr.append(OxmlElement("w:b"))
    if italic:
        r_pr.append(OxmlElement("w:i"))

    col = OxmlElement("w:color")
    col.set(qn("w:val"), color_hex)
    r_pr.append(col)

    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    r_pr.append(u)

    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size_pt * 2)))
    r_pr.append(sz)

    new_run.append(r_pr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def _populate_runs(
    paragraph: Any,
    block: dict[str, Any],
    *,
    font_name: str,
    body_size: float,
    palette: dict[str, str],
    lang: str,
) -> None:
    runs = block.get("runs")
    if isinstance(runs, list) and runs:
        for item in runs:
            txt = apply_french_typography(str(item.get("text", "")), lang=lang)
            if not txt:
                continue
            url = item.get("url")
            bold = bool(item.get("bold", False))
            italic = bool(item.get("italic", False))
            if url:
                _add_hyperlink(
                    paragraph,
                    txt,
                    str(url),
                    font_name=font_name,
                    size_pt=body_size,
                    color_hex=palette["accent"],
                    bold=bold,
                    italic=italic,
                )
            else:
                run = paragraph.add_run(txt)
                run.font.name = font_name
                run.font.size = Pt(body_size)
                run.bold = bold
                run.italic = italic
        return

    raw_text = str(block.get("text", ""))
    formatted = apply_french_typography(raw_text, lang=lang)
    run = paragraph.add_run(formatted)
    run.font.name = font_name
    run.font.size = Pt(body_size)
    if block.get("italic"):
        run.italic = True


def _add_cover_page(
    doc: Document,
    spec: dict[str, Any],
    *,
    base_dir: Path,
    font_name: str,
    palette: dict[str, str],
    lang: str,
) -> None:
    accent_rgb = _hex_to_rgb(palette["accent"])
    text_rgb = _hex_to_rgb(palette["text"])
    muted_rgb = _hex_to_rgb(palette["muted"])

    logo_path_raw = spec.get("logo")
    if logo_path_raw:
        logo_path = _resolve_asset_path(str(logo_path_raw), base_dir)
        w_px, h_px = _get_image_dimensions_px(logo_path)
        ratio = w_px / max(h_px, 1)
        max_w_cm, max_h_cm = 4.5, 2.4
        width_cm = min(max_w_cm, max_h_cm * ratio)
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_logo.paragraph_format.space_before = Pt(0)
        p_logo.paragraph_format.space_after = Pt(14)
        p_logo.add_run().add_picture(str(logo_path), width=Cm(width_cm))

    org = spec.get("organization")
    if org:
        p_org = doc.add_paragraph()
        p_org.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_org.paragraph_format.space_before = Pt(6)
        p_org.paragraph_format.space_after = Pt(6)
        r_org = p_org.add_run(apply_french_typography(str(org), lang=lang))
        r_org.font.name = font_name
        r_org.font.size = Pt(12)
        r_org.font.bold = True
        r_org.font.color.rgb = accent_rgb

    doc_type_label = spec.get("document_type_label")
    if doc_type_label:
        p_type = doc.add_paragraph()
        p_type.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_type.paragraph_format.space_before = Pt(44 if not org else 36)
        p_type.paragraph_format.space_after = Pt(8)
        r_type = p_type.add_run(
            apply_french_typography(str(doc_type_label).upper(), lang=lang)
        )
        r_type.font.name = font_name
        r_type.font.size = Pt(10)
        r_type.font.bold = True
        r_type.font.color.rgb = muted_rgb
    else:
        # Espace vertical avant le titre principal sans paragraphes vides empilés
        pass

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_title.paragraph_format.space_before = Pt(10 if doc_type_label else 68)
    p_title.paragraph_format.space_after = Pt(10)
    p_title.paragraph_format.line_spacing = 1.15
    r_title = p_title.add_run(apply_french_typography(str(spec["title"]), lang=lang))
    r_title.font.name = font_name
    r_title.font.size = Pt(24)
    r_title.font.bold = True
    r_title.font.color.rgb = accent_rgb

    subtitle = spec.get("subtitle")
    if subtitle:
        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_sub.paragraph_format.space_before = Pt(0)
        p_sub.paragraph_format.space_after = Pt(16)
        p_sub.paragraph_format.line_spacing = 1.25
        r_sub = p_sub.add_run(apply_french_typography(str(subtitle), lang=lang))
        r_sub.font.name = font_name
        r_sub.font.size = Pt(13)
        r_sub.font.color.rgb = muted_rgb
        _set_paragraph_bottom_border(
            p_sub, palette["accent"], size_eighths=10, space=12
        )
    else:
        _set_paragraph_bottom_border(
            p_title, palette["accent"], size_eighths=10, space=12
        )

    meta_fields: list[tuple[str, str]] = []
    field_map = [
        ("program", "Filière"),
        ("author", "Auteur"),
        ("supervisor", "Encadrement"),
        ("recipient", "Destinataire"),
        ("reference", "Référence"),
        ("date", "Date"),
        ("academic_year", "Année académique"),
    ]
    for key, label in field_map:
        val = spec.get(key)
        if val:
            meta_fields.append((label, str(val)))

    if meta_fields:
        for idx, (label, val) in enumerate(meta_fields):
            p_meta = doc.add_paragraph()
            p_meta.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p_meta.paragraph_format.space_before = Pt(56 if idx == 0 else 4)
            p_meta.paragraph_format.space_after = Pt(4)
            p_meta.paragraph_format.line_spacing = 1.2
            r_lbl = p_meta.add_run(f"{label}\u00a0: ")
            r_lbl.font.name = font_name
            r_lbl.font.size = Pt(10.5)
            r_lbl.font.bold = True
            r_lbl.font.color.rgb = muted_rgb

            r_val = p_meta.add_run(apply_french_typography(val, lang=lang))
            r_val.font.name = font_name
            r_val.font.size = Pt(10.5)
            r_val.font.color.rgb = text_rgb


def _add_expose_header(
    doc: Document,
    spec: dict[str, Any],
    *,
    font_name: str,
    palette: dict[str, str],
    lang: str,
) -> None:
    accent_rgb = _hex_to_rgb(palette["accent"])
    muted_rgb = _hex_to_rgb(palette["muted"])

    meta_parts = [
        str(spec[k])
        for k in ("organization", "program", "author", "date")
        if spec.get(k)
    ]
    if meta_parts:
        p_meta = doc.add_paragraph()
        p_meta.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_meta.paragraph_format.space_before = Pt(0)
        p_meta.paragraph_format.space_after = Pt(6)
        r_meta = p_meta.add_run(
            apply_french_typography("  |  ".join(meta_parts), lang=lang)
        )
        r_meta.font.name = font_name
        r_meta.font.size = Pt(9.5)
        r_meta.font.color.rgb = muted_rgb

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_title.paragraph_format.space_before = Pt(2)
    p_title.paragraph_format.space_after = Pt(12)
    p_title.paragraph_format.line_spacing = 1.15
    r_title = p_title.add_run(apply_french_typography(str(spec["title"]), lang=lang))
    r_title.font.name = font_name
    r_title.font.size = Pt(18)
    r_title.font.bold = True
    r_title.font.color.rgb = accent_rgb
    _set_paragraph_bottom_border(p_title, palette["accent"], size_eighths=8, space=8)

    if spec.get("subtitle"):
        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_sub.paragraph_format.space_before = Pt(0)
        p_sub.paragraph_format.space_after = Pt(12)
        r_sub = p_sub.add_run(apply_french_typography(str(spec["subtitle"]), lang=lang))
        r_sub.font.name = font_name
        r_sub.font.size = Pt(11.5)
        r_sub.font.color.rgb = muted_rgb


def _add_note_header(
    doc: Document,
    spec: dict[str, Any],
    *,
    font_name: str,
    palette: dict[str, str],
    lang: str,
) -> None:
    accent_rgb = _hex_to_rgb(palette["accent"])
    muted_rgb = _hex_to_rgb(palette["muted"])
    text_rgb = _hex_to_rgb(palette["text"])

    if spec.get("organization"):
        p_org = doc.add_paragraph()
        p_org.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_org.paragraph_format.space_before = Pt(0)
        p_org.paragraph_format.space_after = Pt(8)
        r_org = p_org.add_run(
            apply_french_typography(str(spec["organization"]), lang=lang)
        )
        r_org.font.name = font_name
        r_org.font.size = Pt(11)
        r_org.font.bold = True
        r_org.font.color.rgb = accent_rgb

    header_rows: list[tuple[str, str]] = []
    for key, label in (
        ("author", "Émetteur"),
        ("recipient", "Destinataire"),
        ("date", "Date"),
        ("reference", "Référence"),
        ("subject", "Objet"),
    ):
        val = spec.get(key)
        if not val and key == "subject":
            val = spec.get("title")
        if val:
            header_rows.append((label, str(val)))

    for idx, (label, val) in enumerate(header_rows):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(12 if idx == len(header_rows) - 1 else 3)
        p.paragraph_format.line_spacing = 1.15
        r_lbl = p.add_run(f"{label}\u00a0: ")
        r_lbl.font.name = font_name
        r_lbl.font.size = Pt(10.5)
        r_lbl.font.bold = True
        r_lbl.font.color.rgb = muted_rgb

        r_val = p.add_run(apply_french_typography(val, lang=lang))
        r_val.font.name = font_name
        r_val.font.size = Pt(10.5)
        r_val.font.bold = label == "Objet"
        r_val.font.color.rgb = text_rgb

        if idx == len(header_rows) - 1:
            _set_paragraph_bottom_border(p, palette["accent"], size_eighths=8, space=8)


def _collect_toc_headings(spec: dict[str, Any], lang: str) -> list[tuple[int, str]]:
    headings: list[tuple[int, str]] = []
    for section in spec.get("sections", []):
        level = int(section.get("level", 1))
        level = max(1, min(3, level))
        title = apply_french_typography(str(section.get("title", "")), lang=lang)
        if title:
            headings.append((level, title))
    if spec.get("bibliography"):
        bib_title = apply_french_typography(
            str(spec.get("bibliography_title", "Bibliographie et sources")), lang=lang
        )
        headings.append((1, bib_title))
    return headings


def _add_toc_block(
    doc: Document,
    *,
    toc_title: str,
    headings: list[tuple[int, str]],
    font_name: str,
    palette: dict[str, str],
    lang: str,
) -> None:
    p_heading = doc.add_paragraph()
    p_heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_heading.paragraph_format.space_before = Pt(6)
    p_heading.paragraph_format.space_after = Pt(10)
    p_heading.paragraph_format.keep_with_next = True
    r_head = p_heading.add_run(apply_french_typography(toc_title, lang=lang))
    r_head.font.name = font_name
    r_head.font.size = Pt(15)
    r_head.font.bold = True
    r_head.font.color.rgb = _hex_to_rgb(palette["accent"])

    # Champ Word dynamique TOC avec entrées de prévisualisation entre separate et end
    p_field = doc.add_paragraph()
    p_field.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_field.paragraph_format.space_before = Pt(0)
    p_field.paragraph_format.space_after = Pt(4)

    r_begin = OxmlElement("w:r")
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    fld_begin.set(qn("w:dirty"), "true")
    r_begin.append(fld_begin)

    r_instr = OxmlElement("w:r")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = ' TOC \\o "1-3" \\h \\z \\u '
    r_instr.append(instr)

    r_sep = OxmlElement("w:r")
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    r_sep.append(fld_sep)

    p_field._p.append(r_begin)
    p_field._p.append(r_instr)
    p_field._p.append(r_sep)

    preview_items = headings if headings else [(1, "Document")]
    text_rgb = _hex_to_rgb(palette["text"])
    muted_rgb = _hex_to_rgb(palette["muted"])

    for idx, (level, title) in enumerate(preview_items):
        p_entry = p_field if idx == 0 else doc.add_paragraph()
        if idx > 0:
            p_entry.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p_entry.paragraph_format.space_before = Pt(0)
            p_entry.paragraph_format.space_after = Pt(3)
        p_entry.paragraph_format.left_indent = Cm((level - 1) * 0.65)
        p_entry.paragraph_format.line_spacing = 1.15
        run = p_entry.add_run(title)
        run.font.name = font_name
        run.font.size = Pt(10.5 if level == 1 else 10.0)
        run.font.bold = level == 1
        run.font.color.rgb = text_rgb if level == 1 else muted_rgb

        if idx == len(preview_items) - 1:
            r_end = OxmlElement("w:r")
            fld_end = OxmlElement("w:fldChar")
            fld_end.set(qn("w:fldCharType"), "end")
            r_end.append(fld_end)
            p_entry._p.append(r_end)


def _set_cell_margins(
    cell: Any, top: int = 100, bottom: int = 100, left: int = 140, right: int = 140
) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.find(qn("w:tcMar"))
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for side, val in (
        ("top", top),
        ("bottom", bottom),
        ("left", left),
        ("right", right),
    ):
        el = tc_mar.find(qn(f"w:{side}"))
        if el is None:
            el = OxmlElement(f"w:{side}")
            tc_mar.append(el)
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")


def _set_cell_shading(cell: Any, fill_hex: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(
        f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{fill_hex}"/>'
    )
    tc_pr.append(shd)


def _set_cell_borders(
    cell: Any,
    *,
    top: tuple[str, int] | None = None,
    bottom: tuple[str, int] | None = None,
) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = OxmlElement("w:tcBorders")
    for side in ("left", "right"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "nil")
        tc_borders.append(b)
    for side, spec in (("top", top), ("bottom", bottom)):
        b = OxmlElement(f"w:{side}")
        if spec is None:
            b.set(qn("w:val"), "nil")
        else:
            b.set(qn("w:val"), "single")
            b.set(qn("w:sz"), str(spec[1]))
            b.set(qn("w:space"), "0")
            b.set(qn("w:color"), spec[0])
        tc_borders.append(b)
    tc_pr.append(tc_borders)


def _set_table_horizontal_borders(table: Any, palette: dict[str, str]) -> None:
    tbl_pr = table._tbl.tblPr
    tbl_borders = parse_xml(
        f"<w:tblBorders {nsdecls('w')}>"
        f'<w:top w:val="single" w:sz="6" w:space="0" w:color="{palette["border_strong"]}"/>'
        f'<w:bottom w:val="single" w:sz="8" w:space="0" w:color="{palette["border_strong"]}"/>'
        f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="{palette["border"]}"/>'
        f'<w:left w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'<w:insideV w:val="none"/>'
        f"</w:tblBorders>"
    )
    tbl_pr.append(tbl_borders)


def _render_table_block(
    doc: Document,
    block: dict[str, Any],
    *,
    usable_width_cm: float,
    font_name: str,
    body_size: float,
    palette: dict[str, str],
    lang: str,
) -> None:
    headers = [str(h) for h in block.get("headers", [])]
    rows = block.get("rows", [])
    if not headers and not rows:
        return

    num_cols = len(headers) if headers else len(rows[0])
    caption = block.get("caption")
    if caption:
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_cap.paragraph_format.space_before = Pt(8)
        p_cap.paragraph_format.space_after = Pt(4)
        p_cap.paragraph_format.keep_with_next = True
        r_cap = p_cap.add_run(apply_french_typography(str(caption), lang=lang))
        r_cap.font.name = font_name
        r_cap.font.size = Pt(10)
        r_cap.font.bold = True
        r_cap.font.color.rgb = _hex_to_rgb(palette["accent"])

    total_rows = (1 if headers else 0) + len(rows)
    table = doc.add_table(rows=total_rows, cols=num_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    _set_table_horizontal_borders(table, palette)

    col_widths_raw = block.get("col_widths_cm")
    if isinstance(col_widths_raw, list) and len(col_widths_raw) == num_cols:
        col_widths = [float(w) for w in col_widths_raw]
    else:
        col_widths = [usable_width_cm / num_cols] * num_cols

    alignments_raw = block.get("alignments", [])
    align_map = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }

    text_rgb = _hex_to_rgb(palette["text"])
    muted_rgb = _hex_to_rgb(palette["muted"])
    table_pt = max(9.5, body_size - 1.0)

    row_cursor = 0
    if headers:
        hdr_row = table.rows[0]
        tr_pr = hdr_row._tr.get_or_add_trPr()
        tr_pr.append(OxmlElement("w:tblHeader"))
        tr_pr.append(OxmlElement("w:cantSplit"))
        for col_idx, htext in enumerate(headers):
            cell = hdr_row.cells[col_idx]
            cell.width = Cm(col_widths[col_idx])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            _set_cell_margins(cell, top=110, bottom=110, left=130, right=130)
            _set_cell_shading(cell, palette["table_header"])
            _set_cell_borders(
                cell,
                top=(palette["border_strong"], 6),
                bottom=(palette["border_strong"], 8),
            )
            p = cell.paragraphs[0]
            p.text = ""
            p.alignment = align_map.get(
                alignments_raw[col_idx] if col_idx < len(alignments_raw) else "left",
                WD_ALIGN_PARAGRAPH.LEFT,
            )
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.1
            r = p.add_run(apply_french_typography(htext, lang=lang))
            r.font.name = font_name
            r.font.size = Pt(table_pt)
            r.font.bold = True
            r.font.color.rgb = text_rgb
        row_cursor = 1

    for r_idx, row_vals in enumerate(rows):
        row = table.rows[row_cursor + r_idx]
        tr_pr = row._tr.get_or_add_trPr()
        tr_pr.append(OxmlElement("w:cantSplit"))
        is_last = r_idx == len(rows) - 1
        for col_idx in range(num_cols):
            val = str(row_vals[col_idx]) if col_idx < len(row_vals) else ""
            cell = row.cells[col_idx]
            cell.width = Cm(col_widths[col_idx])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
            _set_cell_margins(cell, top=95, bottom=95, left=130, right=130)
            _set_cell_borders(
                cell,
                bottom=(palette["border_strong"], 8)
                if is_last
                else (palette["border"], 4),
            )
            p = cell.paragraphs[0]
            p.text = ""
            p.alignment = align_map.get(
                alignments_raw[col_idx] if col_idx < len(alignments_raw) else "left",
                WD_ALIGN_PARAGRAPH.LEFT,
            )
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.12
            r = p.add_run(apply_french_typography(val, lang=lang))
            r.font.name = font_name
            r.font.size = Pt(table_pt)
            r.font.color.rgb = text_rgb

    source = block.get("source")
    if source:
        p_src = doc.add_paragraph()
        p_src.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_src.paragraph_format.space_before = Pt(4)
        p_src.paragraph_format.space_after = Pt(10)
        r_src = p_src.add_run(apply_french_typography(f"Source : {source}", lang=lang))
        r_src.font.name = font_name
        r_src.font.size = Pt(9)
        r_src.font.italic = True
        r_src.font.color.rgb = muted_rgb
    else:
        p_after = doc.add_paragraph()
        p_after.paragraph_format.space_before = Pt(0)
        p_after.paragraph_format.space_after = Pt(4)
        p_after.paragraph_format.line_spacing = 1.0


def _render_blocks(
    doc: Document,
    blocks: list[dict[str, Any]],
    *,
    base_dir: Path,
    usable_width_cm: float,
    font_name: str,
    body_size: float,
    palette: dict[str, str],
    lang: str,
) -> None:
    for block in blocks:
        btype = str(block.get("type", "paragraph")).lower()

        if btype == "paragraph":
            p = doc.add_paragraph(style="Normal")
            _populate_runs(
                p,
                block,
                font_name=font_name,
                body_size=body_size,
                palette=palette,
                lang=lang,
            )
        elif btype in {"bullets", "numbered"}:
            style_name = "List Bullet" if btype == "bullets" else "List Number"
            for item in block.get("items", []):
                p = doc.add_paragraph(style=style_name)
                if isinstance(item, dict):
                    _populate_runs(
                        p,
                        item,
                        font_name=font_name,
                        body_size=body_size,
                        palette=palette,
                        lang=lang,
                    )
                else:
                    r = p.add_run(apply_french_typography(str(item), lang=lang))
                    r.font.name = font_name
                    r.font.size = Pt(body_size)
        elif btype == "quote":
            p = doc.add_paragraph(style="Normal")
            p.paragraph_format.left_indent = Cm(0.9)
            p.paragraph_format.right_indent = Cm(0.6)
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(8)
            _set_paragraph_left_border(p, palette["accent"], size_eighths=16, space=10)
            txt = apply_french_typography(str(block.get("text", "")), lang=lang)
            r = p.add_run(txt)
            r.font.name = font_name
            r.font.size = Pt(max(10.0, body_size - 0.5))
            r.font.italic = True
            r.font.color.rgb = _hex_to_rgb(palette["muted"])
        elif btype == "table":
            _render_table_block(
                doc,
                block,
                usable_width_cm=usable_width_cm,
                font_name=font_name,
                body_size=body_size,
                palette=palette,
                lang=lang,
            )
        elif btype == "image":
            img_path = _resolve_asset_path(str(block["path"]), base_dir)
            width_cm = min(
                usable_width_cm, float(block.get("width_cm", usable_width_cm * 0.85))
            )
            caption = block.get("caption")
            if caption:
                p_cap = doc.add_paragraph()
                p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_cap.paragraph_format.space_before = Pt(8)
                p_cap.paragraph_format.space_after = Pt(4)
                p_cap.paragraph_format.keep_with_next = True
                r_cap = p_cap.add_run(apply_french_typography(str(caption), lang=lang))
                r_cap.font.name = font_name
                r_cap.font.size = Pt(9.5)
                r_cap.font.bold = True
                r_cap.font.color.rgb = _hex_to_rgb(palette["muted"])
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(4)
            p_img.paragraph_format.space_after = Pt(8)
            p_img.add_run().add_picture(str(img_path), width=Cm(width_cm))
        elif btype == "page_break":
            doc.add_page_break()
        else:
            raise ValueError(f"Type de bloc inconnu : {btype!r}")


def _render_bibliography(
    doc: Document,
    entries: list[Any],
    *,
    title: str,
    font_name: str,
    body_size: float,
    palette: dict[str, str],
    lang: str,
) -> None:
    doc.add_heading(apply_french_typography(title, lang=lang), level=1)
    bib_size = max(10.0, body_size - 0.5)

    for entry in entries:
        p = doc.add_paragraph(style="Normal")
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.left_indent = Cm(1.25)
        p.paragraph_format.first_line_indent = Cm(-1.25)
        p.paragraph_format.space_after = Pt(5)
        p.paragraph_format.line_spacing = 1.15

        if isinstance(entry, str):
            r = p.add_run(apply_french_typography(entry, lang=lang))
            r.font.name = font_name
            r.font.size = Pt(bib_size)
        elif isinstance(entry, dict):
            citation = str(entry.get("citation", "")).strip()
            url = entry.get("url")
            if citation:
                r = p.add_run(
                    apply_french_typography(citation + (" " if url else ""), lang=lang)
                )
                r.font.name = font_name
                r.font.size = Pt(bib_size)
            if url:
                cleaned_url, _ = clean_url(str(url))
                _add_hyperlink(
                    p,
                    cleaned_url,
                    cleaned_url,
                    font_name=font_name,
                    size_pt=bib_size,
                    color_hex=palette["accent"],
                )


def _default_output_path(spec: dict[str, Any], genre: str, out_dir: Path) -> Path:
    prefix = GENRE_DEFAULTS[genre]["prefix"]
    raw_title = str(spec["title"])
    short = " ".join(raw_title.split()[:6])
    name_parts = [prefix, sanitize_filename(short)]
    author = spec.get("author")
    if genre == "memoire" and author:
        last_token = str(author).strip().split()[-1]
        name_parts.append(sanitize_filename(last_token))
    filename = " - ".join(part for part in name_parts if part) + ".docx"
    return out_dir / filename


def build_docx(
    spec: dict[str, Any],
    output_path: Path | None = None,
    *,
    base_dir: Path | None = None,
    allow_placeholders: bool = False,
    allow_custom_color: bool = False,
) -> Path:
    """Construit un fichier .docx complet et nettoyé à partir d'un dictionnaire JSON."""
    if "title" not in spec or not str(spec["title"]).strip():
        raise ValueError("Le champ obligatoire 'title' est manquant ou vide.")

    genre = str(spec.get("genre", "rapport")).strip().lower()
    if genre not in GENRE_DEFAULTS:
        raise ValueError(
            f"Genre invalide : {genre!r}. Choix valides : {', '.join(GENRE_DEFAULTS)}"
        )

    if not allow_placeholders:
        hits = _scan_placeholders(spec)
        if hits:
            raise ValueError(
                "Texte factice détecté dans le JSON : " + "; ".join(hits[:5])
            )

    defaults = GENRE_DEFAULTS[genre]
    lang = str(spec.get("lang", "fr")).strip().lower()
    font_name = str(spec.get("font") or defaults["font"]).strip()
    if font_name not in NATIVE_FONTS:
        raise ValueError(
            f"Police non native refusée : {font_name!r}. "
            f"Polices autorisées : {', '.join(sorted(NATIVE_FONTS))}"
        )

    body_size = float(spec.get("body_size_pt") or defaults["body_size"])
    line_spacing = float(spec.get("line_spacing") or defaults["line_spacing"])
    margins_cm = dict(defaults["margins_cm"])
    if isinstance(spec.get("margins_cm"), dict):
        margins_cm.update(spec["margins_cm"])

    palette = _resolve_palette(
        spec.get("palette"), defaults["palette"], allow_custom_color
    )
    cover_page = bool(spec.get("cover_page", defaults["cover_page"]))
    toc_enabled = bool(spec.get("toc", defaults["toc"]))
    toc_title = str(spec.get("toc_title") or defaults["toc_title"])
    show_header = bool(spec.get("header", defaults["header"]))
    show_page_numbers = bool(spec.get("page_numbers", defaults["page_numbers"]))
    short_title = apply_french_typography(
        str(spec.get("short_title") or spec["title"]), lang=lang
    )

    resolved_base_dir = base_dir or Path.cwd()
    dest = output_path or _default_output_path(spec, genre, resolved_base_dir)
    dest = dest.resolve()
    dest.parent.mkdir(parents=True, exist_ok=True)

    doc = Document()
    _enable_update_fields_on_open(doc)
    _configure_styles(
        doc,
        font_name=font_name,
        body_size=body_size,
        line_spacing=line_spacing,
        palette=palette,
        genre=genre,
        lang=lang,
    )

    sec1 = doc.sections[0]
    sec1.page_width = Cm(21.0)
    sec1.page_height = Cm(29.7)
    _set_section_margins(sec1, margins_cm)

    usable_width_cm = 21.0 - float(margins_cm["left"]) - float(margins_cm["right"])
    preliminaries = spec.get("preliminaries", [])
    has_prelims = bool(
        cover_page
        and genre == "memoire"
        and isinstance(preliminaries, list)
        and preliminaries
    )

    if cover_page:
        _configure_headers_and_footers(
            sec1,
            show_header=show_header and not has_prelims,
            show_page_numbers=show_page_numbers,
            different_first_page=True,
            short_title=short_title,
            font_name=font_name,
            palette=palette,
        )
        if has_prelims:
            _set_section_page_number_format(sec1, fmt="lowerRoman", start=1)
        _add_cover_page(
            doc,
            spec,
            base_dir=resolved_base_dir,
            font_name=font_name,
            palette=palette,
            lang=lang,
        )
        doc.add_page_break()
    elif genre == "expose":
        _configure_headers_and_footers(
            sec1,
            show_header=show_header,
            show_page_numbers=show_page_numbers,
            different_first_page=True,
            short_title=short_title,
            font_name=font_name,
            palette=palette,
        )
        _add_expose_header(doc, spec, font_name=font_name, palette=palette, lang=lang)
    elif genre == "note":
        _configure_headers_and_footers(
            sec1,
            show_header=show_header,
            show_page_numbers=show_page_numbers,
            different_first_page=False,
            short_title=short_title,
            font_name=font_name,
            palette=palette,
        )
        _add_note_header(doc, spec, font_name=font_name, palette=palette, lang=lang)

    if has_prelims:
        for idx, prelim in enumerate(preliminaries):
            ptitle = prelim.get("title")
            if ptitle:
                p_h = doc.add_paragraph()
                p_h.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p_h.paragraph_format.space_before = Pt(6)
                p_h.paragraph_format.space_after = Pt(10)
                p_h.paragraph_format.keep_with_next = True
                r_ph = p_h.add_run(apply_french_typography(str(ptitle), lang=lang))
                r_ph.font.name = font_name
                r_ph.font.size = Pt(15)
                r_ph.font.bold = True
                r_ph.font.color.rgb = _hex_to_rgb(palette["accent"])
            _render_blocks(
                doc,
                prelim.get("blocks", []),
                base_dir=resolved_base_dir,
                usable_width_cm=usable_width_cm,
                font_name=font_name,
                body_size=body_size,
                palette=palette,
                lang=lang,
            )
            if idx < len(preliminaries) - 1:
                doc.add_page_break()

        sec2 = doc.add_section(WD_SECTION.NEW_PAGE)
        sec2.page_width = Cm(21.0)
        sec2.page_height = Cm(29.7)
        _set_section_margins(sec2, margins_cm)
        _set_section_page_number_format(sec2, fmt="decimal", start=1)
        _configure_headers_and_footers(
            sec2,
            show_header=show_header,
            show_page_numbers=show_page_numbers,
            different_first_page=False,
            short_title=short_title,
            font_name=font_name,
            palette=palette,
        )

    just_broke_page = cover_page
    if toc_enabled:
        headings = _collect_toc_headings(spec, lang=lang)
        _add_toc_block(
            doc,
            toc_title=toc_title,
            headings=headings,
            font_name=font_name,
            palette=palette,
            lang=lang,
        )
        doc.add_page_break()
        just_broke_page = True

    for section in spec.get("sections", []):
        level = max(1, min(3, int(section.get("level", 1))))
        stitle = apply_french_typography(str(section.get("title", "")), lang=lang)
        if stitle:
            if section.get("page_break_before") and not just_broke_page:
                doc.add_page_break()
            doc.add_heading(stitle, level=level)
            just_broke_page = False
        blocks = section.get("blocks", [])
        if blocks:
            _render_blocks(
                doc,
                blocks,
                base_dir=resolved_base_dir,
                usable_width_cm=usable_width_cm,
                font_name=font_name,
                body_size=body_size,
                palette=palette,
                lang=lang,
            )
            just_broke_page = str(blocks[-1].get("type", "")).lower() == "page_break"

    bib_entries = spec.get("bibliography", [])
    if isinstance(bib_entries, list) and bib_entries:
        _render_bibliography(
            doc,
            bib_entries,
            title=str(spec.get("bibliography_title", "Bibliographie et sources")),
            font_name=font_name,
            body_size=body_size,
            palette=palette,
            lang=lang,
        )

    doc.save(str(dest))
    sanitize_docx(
        dest,
        author=str(spec.get("author", "")),
        title=apply_french_typography(str(spec["title"]), lang=lang),
        in_place=True,
    )
    return dest


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Construit un fichier .docx officiel depuis une spécification JSON."
    )
    parser.add_argument("spec_json", type=Path, help="Chemin du fichier JSON source.")
    parser.add_argument(
        "-o", "--output", type=Path, default=None, help="Chemin .docx cible."
    )
    parser.add_argument(
        "--allow-placeholders",
        action="store_true",
        help="Autorise exceptionnellement les balises factices.",
    )
    parser.add_argument(
        "--allow-custom-color",
        action="store_true",
        help="Autorise une couleur hexadécimale hors palettes recommandées.",
    )
    args = parser.parse_args(argv)

    spec_path = args.spec_json.resolve()
    spec = json.loads(spec_path.read_text(encoding="utf-8"))
    out_path = build_docx(
        spec,
        output_path=args.output,
        base_dir=spec_path.parent,
        allow_placeholders=args.allow_placeholders,
        allow_custom_color=args.allow_custom_color,
    )
    print(f"Document généré et assaini : {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

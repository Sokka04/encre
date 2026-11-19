# /// script
# requires-python = ">=3.11"
# dependencies = [
#   "pymupdf>=1.24.0",
# ]
# ///
"""Convertit un .docx en PDF puis en images PNG page par page pour inspection visuelle."""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def _find_soffice() -> Path | None:
    on_path = shutil.which("soffice") or shutil.which("libreoffice")
    if on_path:
        return Path(on_path)
    candidates = [
        Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
        Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
        Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
    ]
    for c in candidates:
        if c.exists():
            return c
    return None


def _has_ms_word() -> bool:
    if sys.platform != "win32":
        return False
    for c in (
        Path(r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"),
        Path(r"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE"),
        Path(r"C:\Program Files\Microsoft Office\Office16\WINWORD.EXE"),
    ):
        if c.exists():
            return True
    try:
        import winreg

        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE",
        ) as key:
            val, _ = winreg.QueryValueEx(key, "")
            return bool(val and Path(val).exists())
    except OSError:
        return False


def _convert_docx_to_pdf_soffice(soffice: Path, docx_path: Path, out_pdf: Path) -> None:
    cmd = [
        str(soffice),
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        str(out_pdf.parent),
        str(docx_path),
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True, check=False)
    if proc.returncode != 0:
        raise RuntimeError(
            f"Échec LibreOffice ({proc.returncode}) : {proc.stderr or proc.stdout}"
        )
    produced = out_pdf.parent / f"{docx_path.stem}.pdf"
    if not produced.exists():
        raise RuntimeError(
            f"PDF attendu introuvable après conversion LibreOffice : {produced}"
        )
    if produced.resolve() != out_pdf.resolve():
        shutil.move(str(produced), str(out_pdf))


def _convert_docx_to_pdf_word_com(docx_path: Path, out_pdf: Path) -> None:
    ps_script = r"""
$ErrorActionPreference = 'Stop'
$inPath = $env:ENCRE_IN_DOCX
$outPath = $env:ENCRE_OUT_PDF
$word = $null
$doc = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($inPath, $false, $true)
    foreach ($toc in $doc.TablesOfContents) {
        $toc.Update()
    }
    $doc.Saved = $true
    $doc.ExportAsFixedFormat($outPath, 17)
}
finally {
    if ($null -ne $doc) {
        try {
            $doc.Saved = $true
            $doc.Close([ref]$false)
        } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
    }
    if ($null -ne $word) {
        try {
            $word.Quit([ref]$false)
        } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
}
"""
    env = os.environ.copy()
    env["ENCRE_IN_DOCX"] = str(docx_path)
    env["ENCRE_OUT_PDF"] = str(out_pdf)

    proc = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_script],
        env=env,
        capture_output=True,
        text=True,
        check=False,
    )
    if proc.returncode != 0 or not out_pdf.exists():
        raise RuntimeError(
            f"Échec export PDF via Word COM ({proc.returncode}) : {proc.stderr.strip() or proc.stdout.strip()}"
        )


def _compute_white_ratio_rgb(samples: bytes, stride_pixels: int = 4) -> float:
    step = stride_pixels * 3
    total = 0
    white = 0
    mv = memoryview(samples)
    for idx in range(0, len(mv) - 2, step):
        total += 1
        if mv[idx] >= 250 and mv[idx + 1] >= 250 and mv[idx + 2] >= 250:
            white += 1
    return (white / total) if total else 1.0


def _render_pdf_with_fitz(
    pdf_path: Path, out_dir: Path, dpi: int
) -> list[dict[str, Any]]:
    try:
        import pymupdf  # type: ignore[import-not-found]
    except ImportError:
        import fitz as pymupdf  # type: ignore[import-not-found]

    pages_info: list[dict[str, Any]] = []
    zoom = dpi / 72.0
    matrix = pymupdf.Matrix(zoom, zoom)

    with pymupdf.open(str(pdf_path)) as pdf:
        for idx, page in enumerate(pdf, start=1):
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            out_png = out_dir / f"page-{idx:02d}.png"
            pix.save(str(out_png))
            white_ratio = _compute_white_ratio_rgb(pix.samples)
            pages_info.append(
                {
                    "page": idx,
                    "path": str(out_png),
                    "width": pix.width,
                    "height": pix.height,
                    "white_ratio": round(white_ratio, 4),
                    "blank_suspect": white_ratio >= 0.995,
                }
            )
    return pages_info


def _render_pdf_with_pdftoppm(
    pdf_path: Path, out_dir: Path, dpi: int
) -> list[dict[str, Any]]:
    pdftoppm = shutil.which("pdftoppm")
    if not pdftoppm:
        raise RuntimeError(
            "Aucun moteur de rendu PDF->PNG disponible (installez pymupdf ou pdftoppm)."
        )
    prefix = out_dir / "_raw_page"
    proc = subprocess.run(
        [pdftoppm, "-png", "-r", str(dpi), str(pdf_path), str(prefix)],
        capture_output=True,
        text=True,
        check=False,
    )
    if proc.returncode != 0:
        raise RuntimeError(
            f"Échec pdftoppm ({proc.returncode}) : {proc.stderr or proc.stdout}"
        )

    generated = sorted(out_dir.glob("_raw_page-*.png"))
    pages_info: list[dict[str, Any]] = []
    for idx, src_png in enumerate(generated, start=1):
        dest_png = out_dir / f"page-{idx:02d}.png"
        shutil.move(str(src_png), str(dest_png))
        pages_info.append(
            {
                "page": idx,
                "path": str(dest_png),
                "width": None,
                "height": None,
                "white_ratio": None,
                "blank_suspect": False,
            }
        )
    return pages_info


def render_preview(
    input_docx: Path,
    *,
    out_dir: Path | None = None,
    dpi: int = 150,
    keep_pdf: bool = False,
) -> dict[str, Any]:
    """Convertit un .docx en PDF puis rend chaque page en PNG."""
    src = Path(input_docx).resolve()
    if not src.exists():
        raise FileNotFoundError(f"Fichier introuvable : {src}")

    target_dir = (out_dir or src.with_name(f"{src.stem}-preview")).resolve()
    target_dir.mkdir(parents=True, exist_ok=True)

    soffice = _find_soffice()
    pdf_backend = ""
    if soffice is not None:
        pdf_backend = "libreoffice"
    elif _has_ms_word():
        pdf_backend = "word_com"
    else:
        raise OSError(
            "Aucun moteur PDF disponible (ni LibreOffice soffice, ni Microsoft Word sous Windows)."
        )

    if keep_pdf:
        pdf_path = target_dir / f"{src.stem}.pdf"
        cleanup_pdf = False
    else:
        fd, tmp_pdf = tempfile.mkstemp(
            prefix=".encre_", suffix=".pdf", dir=str(target_dir)
        )
        os.close(fd)
        pdf_path = Path(tmp_pdf)
        cleanup_pdf = True

    try:
        if pdf_backend == "libreoffice" and soffice is not None:
            _convert_docx_to_pdf_soffice(soffice, src, pdf_path)
        else:
            _convert_docx_to_pdf_word_com(src, pdf_path)

        try:
            pages = _render_pdf_with_fitz(pdf_path, target_dir, dpi)
            png_backend = "pymupdf"
        except ImportError:
            pages = _render_pdf_with_pdftoppm(pdf_path, target_dir, dpi)
            png_backend = "pdftoppm"
    finally:
        if cleanup_pdf and pdf_path.exists():
            pdf_path.unlink(missing_ok=True)

    blank_pages = [p["page"] for p in pages if p.get("blank_suspect")]
    return {
        "input": str(src),
        "out_dir": str(target_dir),
        "pdf_path": str(pdf_path) if keep_pdf else None,
        "pdf_backend": pdf_backend,
        "png_backend": png_backend,
        "page_count": len(pages),
        "blank_suspect_pages": blank_pages,
        "pages": pages,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Génère les aperçus PNG page par page d'un fichier .docx."
    )
    parser.add_argument("input_docx", type=Path, help="Fichier .docx à convertir.")
    parser.add_argument(
        "--out-dir", type=Path, default=None, help="Dossier de sortie des PNG."
    )
    parser.add_argument(
        "--dpi", type=int, default=150, help="Résolution de rendu PNG (défaut 150)."
    )
    parser.add_argument(
        "--keep-pdf", action="store_true", help="Conserve le PDF intermédiaire."
    )
    parser.add_argument("--json", action="store_true", help="Sortie JSON structurée.")
    args = parser.parse_args(argv)

    try:
        result = render_preview(
            args.input_docx,
            out_dir=args.out_dir,
            dpi=args.dpi,
            keep_pdf=args.keep_pdf,
        )
    except OSError as exc:
        print(f"Prérequis manquant : {exc}", file=sys.stderr)
        return 2
    except RuntimeError as exc:
        print(f"Erreur rendu : {exc}", file=sys.stderr)
        return 2

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(
            f"Rendu terminé : {result['page_count']} page(s) dans {result['out_dir']} "
            f"(PDF : {result['pdf_backend']}, PNG : {result['png_backend']})"
        )
        for p in result["pages"]:
            flag = " [ATTENTION : page quasi blanche]" if p["blank_suspect"] else ""
            print(f"  - Page {p['page']:02d} : {p['path']}{flag}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

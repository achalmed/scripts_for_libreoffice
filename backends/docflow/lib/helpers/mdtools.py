#!/usr/bin/env python3
"""mdtools.py — Post-procesado y verificación de Markdown de docflow.

Subcomandos:
    fix-media <md> --media-dir DIR   Renombra imágenes a figure-NNN, las mueve
                                     a DIR/images/ y reescribe los enlaces
                                     como rutas relativas al .md.
    retarget-media <md> --media-dir DIR
                                     Reapunta enlaces cuya imagen cambió de
                                     extensión (p. ej. tras convertir a WebP).
    clean <md>                       Limpieza conservadora: HTML residual,
                                     atributos basura, líneas en blanco
                                     duplicadas, normalización UTF-8 (NFC).
    verify <md>                      Verifica contenido no vacío e imágenes
                                     existentes. Sale con 1 si algo falla.
"""

import argparse
import re
import shutil
import sys
import unicodedata
from pathlib import Path
from urllib.parse import quote, unquote

# ![alt](ruta "título opcional")  — la ruta no puede contener ')' sin escapar
IMG_RE = re.compile(r"!\[(?P<alt>[^\]]*)\]\((?P<path>[^)]+?)(?:\s+\"[^\"]*\")?\)")

# pandoc (gfm) emite figuras como HTML crudo: <figure><img …><figcaption>…
FIGURE_RE = re.compile(
    r"<figure>\s*<img\s+[^>]*?src=\"(?P<src>[^\"]+)\"[^>]*/?>\s*"
    r"(?:<figcaption[^>]*>(?P<cap>.*?)</figcaption>)?\s*</figure>",
    re.S,
)
IMG_TAG_RE = re.compile(r"<img\s+[^>]*?src=\"(?P<src>[^\"]+)\"[^>]*/?>")


def html_images_to_md(text: str) -> str:
    """Reduce <figure>/<img> crudos de pandoc a imágenes Markdown puras."""
    def fig(match: re.Match) -> str:
        caption = re.sub(r"<[^>]+>", "", match.group("cap") or "").strip()
        return f"![{caption}]({match.group('src')})"

    text = FIGURE_RE.sub(fig, text)
    return IMG_TAG_RE.sub(lambda m: f"![]({m.group('src')})", text)


def _iter_image_links(text: str):
    return IMG_RE.finditer(text)


def _is_remote(path: str) -> bool:
    return path.startswith(("http://", "https://", "data:", "//"))


# ---------------------------------------------------------------------------
# fix-media
# ---------------------------------------------------------------------------
def fix_media(md_file: Path, media_dir: Path) -> int:
    """Consolida las imágenes en media_dir/images/figure-NNN.* y arregla enlaces."""
    md_file = md_file.resolve()
    media_dir = media_dir.resolve()
    text = html_images_to_md(md_file.read_text(encoding="utf-8", errors="replace"))
    images_dir = media_dir / "images"
    md_dir = md_file.parent

    counter = 0
    mapping: dict[str, str] = {}

    def resolve_source(raw: str) -> Path | None:
        p = Path(unquote(raw))
        candidates = [p] if p.is_absolute() else [md_dir / p, media_dir / p]
        for cand in candidates:
            if cand.is_file():
                return cand
        # pandoc a veces escribe solo el nombre: buscarlo dentro de media_dir
        hits = list(media_dir.rglob(p.name)) if media_dir.is_dir() else []
        return hits[0] if hits else None

    def replace(match: re.Match) -> str:
        nonlocal counter
        raw = match.group("path")
        if _is_remote(raw):
            return match.group(0)
        if raw in mapping:
            return f"![{match.group('alt')}]({mapping[raw]})"
        source = resolve_source(raw)
        if source is None:
            return match.group(0)  # enlace roto: se detectará en verify
        counter += 1
        dest = images_dir / f"figure-{counter:03d}{source.suffix.lower()}"
        dest.parent.mkdir(parents=True, exist_ok=True)
        if source.resolve() != dest.resolve():
            shutil.copy2(source, dest)
        link = quote(str(dest.relative_to(md_dir)))
        mapping[raw] = link
        return f"![{match.group('alt')}]({link})"

    new_text = IMG_RE.sub(replace, text)
    md_file.write_text(new_text, encoding="utf-8")

    # Retirar del media_dir los restos de extracción (dir media/ de pandoc, etc.)
    if media_dir.is_dir():
        for child in list(media_dir.iterdir()):
            if child.name != "images":
                shutil.rmtree(child) if child.is_dir() else child.unlink()
        has_images = images_dir.is_dir() and any(images_dir.iterdir())
        if not has_images:
            shutil.rmtree(media_dir, ignore_errors=True)
    return 0


# ---------------------------------------------------------------------------
# retarget-media
# ---------------------------------------------------------------------------
def retarget_media(md_file: Path, media_dir: Path) -> int:
    """Si una imagen enlazada ya no existe pero sí con otra extensión, reapunta."""
    text = md_file.read_text(encoding="utf-8", errors="replace")
    md_dir = md_file.parent.resolve()

    def replace(match: re.Match) -> str:
        raw = match.group("path")
        if _is_remote(raw):
            return match.group(0)
        target = md_dir / Path(unquote(raw))
        if target.is_file():
            return match.group(0)
        hits = list(target.parent.glob(target.stem + ".*")) if target.parent.is_dir() else []
        if hits:
            link = quote(str(hits[0].relative_to(md_dir)))
            return f"![{match.group('alt')}]({link})"
        return match.group(0)

    md_file.write_text(IMG_RE.sub(replace, text), encoding="utf-8")
    return 0


# ---------------------------------------------------------------------------
# clean
# ---------------------------------------------------------------------------
_CLEAN_PATTERNS = [
    (re.compile(r"<a\s+(?:id|name)=\"[^\"]*\"\s*>\s*</a>"), ""),       # anclas vacías
    (re.compile(r"<span\s+id=\"[^\"]*\"\s*>\s*</span>"), ""),          # spans-ID vacíos
    (re.compile(r"</?div[^>]*>"), ""),                                  # divs sueltos
    (re.compile(r"\{width=\"[^\"]*\"(?:\s+height=\"[^\"]*\")?\}"), ""),  # atributos de tamaño
    (re.compile(r"\{#[\w:-]+(?:\s+\.[\w-]+)*\}"), ""),                  # IDs/clases pandoc
    (re.compile(r"\[\]\{[^}]*\}"), ""),                                 # spans vacíos con atributos
    (re.compile(r"^\s*<!--.*?-->\s*$", re.M | re.S), ""),               # comentarios solitarios
]


def clean(md_file: Path) -> int:
    text = md_file.read_text(encoding="utf-8", errors="replace")
    text = text.lstrip("﻿")                      # BOM
    text = unicodedata.normalize("NFC", text)         # acentos compuestos → NFC

    # Proteger bloques de código de la limpieza
    chunks = re.split(r"(```.*?```|~~~.*?~~~)", text, flags=re.S)
    for i, chunk in enumerate(chunks):
        if chunk.startswith(("```", "~~~")):
            continue
        for pattern, repl in _CLEAN_PATTERNS:
            chunk = pattern.sub(repl, chunk)
        # <img> HTML → imagen Markdown
        chunk = re.sub(
            r"<img\s+[^>]*src=\"([^\"]+)\"[^>]*/?>",
            lambda m: f"![]({m.group(1)})", chunk)
        chunks[i] = chunk
    text = "".join(chunks)

    # Espacios finales de línea y líneas en blanco múltiples
    text = re.sub(r"[ \t]+$", "", text, flags=re.M)
    text = re.sub(r"\n{3,}", "\n\n", text)
    # Garantizar línea en blanco antes de encabezados (legibilidad + parsers)
    text = re.sub(r"(?<=[^\n])\n(#{1,6} )", r"\n\n\1", text)

    md_file.write_text(text.strip() + "\n", encoding="utf-8")
    return 0


# ---------------------------------------------------------------------------
# verify
# ---------------------------------------------------------------------------
def verify(md_file: Path) -> int:
    if not md_file.is_file():
        print(f"verify: no existe: {md_file}", file=sys.stderr)
        return 1
    text = md_file.read_text(encoding="utf-8", errors="replace")

    # Contenido real más allá del frontmatter
    body = re.sub(r"\A---\n.*?\n---\n", "", text, flags=re.S)
    if not body.strip():
        print(f"verify: Markdown sin contenido: {md_file}", file=sys.stderr)
        return 1

    md_dir = md_file.parent
    missing = []
    for match in _iter_image_links(text):
        raw = match.group("path")
        if _is_remote(raw):
            continue
        target = Path(unquote(raw))
        if not (target if target.is_absolute() else md_dir / target).is_file():
            missing.append(raw)
    if missing:
        for m in missing:
            print(f"verify: imagen no encontrada: {m}", file=sys.stderr)
        return 1
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("action", choices=["fix-media", "retarget-media", "clean", "verify"])
    ap.add_argument("md_file", type=Path)
    ap.add_argument("--media-dir", type=Path, default=None)
    args = ap.parse_args()

    if args.action == "fix-media":
        return fix_media(args.md_file, args.media_dir or args.md_file.with_name(args.md_file.stem + "_files"))
    if args.action == "retarget-media":
        return retarget_media(args.md_file, args.media_dir or args.md_file.parent)
    if args.action == "clean":
        return clean(args.md_file)
    return verify(args.md_file)


if __name__ == "__main__":
    sys.exit(main())

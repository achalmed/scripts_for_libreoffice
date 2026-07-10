#!/usr/bin/env python3
"""pptx_to_md.py — Conversor PPTX → Markdown de docflow.

Pandoc no puede leer .pptx, así que docflow parsea directamente el contenedor
OOXML (zip + XML, solo stdlib): una sección por diapositiva, con títulos,
listas anidadas, tablas, hipervínculos, imágenes extraídas y notas del
presentador.

Uso:
    pptx_to_md.py entrada.pptx salida.md --media-dir salida_files
                  [--slide-separator ---] [--no-notes]
"""

import argparse
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from urllib.parse import quote

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def parse_rels(zf: zipfile.ZipFile, part: str) -> dict:
    """Mapa rId → target del archivo .rels asociado a una parte del paquete."""
    rels_path = f"{Path(part).parent}/_rels/{Path(part).name}.rels"
    try:
        root = ET.fromstring(zf.read(rels_path))
    except KeyError:
        return {}
    rels = {}
    for rel in root.findall("rel:Relationship", NS):
        target = rel.get("Target", "")
        if not target.startswith("/") and rel.get("TargetMode", "Internal") == "Internal":
            target = _normalize(Path(part).parent, target)
        rels[rel.get("Id")] = {"target": target.lstrip("/"), "type": rel.get("Type", ""),
                               "mode": rel.get("TargetMode", "Internal")}
    return rels


def _normalize(base: Path, target: str) -> str:
    parts = list(base.parts) + target.split("/")
    out = []
    for p in parts:
        if p == "..":
            out.pop()
        elif p != ".":
            out.append(p)
    return "/".join(out)


def slide_parts_in_order(zf: zipfile.ZipFile) -> list:
    """Diapositivas en el orden real de la presentación (sldIdLst + rels)."""
    rels = parse_rels(zf, "ppt/presentation.xml")
    root = ET.fromstring(zf.read("ppt/presentation.xml"))
    parts = []
    for sld in root.iter(f"{{{NS['p']}}}sldId"):
        rid = sld.get(f"{{{NS['r']}}}id")
        if rid in rels:
            parts.append(rels[rid]["target"])
    if not parts:  # fallback: orden numérico de los archivos
        parts = sorted(
            (n for n in zf.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
            key=lambda n: int(re.search(r"\d+", Path(n).name).group()),
        )
    return parts


def run_text(run: ET.Element, rels: dict) -> str:
    """Texto de un a:r con negrita/cursiva e hipervínculos."""
    text = "".join(t.text or "" for t in run.findall("a:t", NS))
    if not text.strip():
        return text
    props = run.find("a:rPr", NS)
    if props is not None:
        link = props.find("a:hlinkClick", NS)
        if link is not None:
            rid = link.get(f"{{{NS['r']}}}id")
            info = rels.get(rid)
            if info and info["mode"] == "External":
                text = f"[{text}]({info['target']})"
        if props.get("b") == "1":
            text = f"**{text}**"
        elif props.get("i") == "1":
            text = f"*{text}*"
    return text


def paragraph_md(par: ET.Element, rels: dict, as_bullets: bool) -> str:
    """Un a:p como línea Markdown (con nivel de sangría de lista)."""
    text = "".join(run_text(r, rels) for r in par.findall("a:r", NS)).strip()
    if not text:
        return ""
    if not as_bullets:
        return text
    ppr = par.find("a:pPr", NS)
    level = int(ppr.get("lvl", "0")) if ppr is not None else 0
    no_bullet = ppr is not None and ppr.find("a:buNone", NS) is not None
    if no_bullet and level == 0:
        return text
    return f"{'  ' * level}- {text}"


def table_md(tbl: ET.Element, rels: dict) -> str:
    rows = []
    for tr in tbl.findall("a:tr", NS):
        cells = []
        for tc in tr.findall("a:tc", NS):
            parts = [paragraph_md(p, rels, as_bullets=False)
                     for p in tc.iter(f"{{{NS['a']}}}p")]
            cells.append(" ".join(x for x in parts if x).replace("|", "\\|") or " ")
        rows.append(cells)
    if not rows:
        return ""
    width = max(len(r) for r in rows)
    rows = [r + [" "] * (width - len(r)) for r in rows]
    out = ["| " + " | ".join(rows[0]) + " |",
           "|" + "---|" * width]
    out += ["| " + " | ".join(r) + " |" for r in rows[1:]]
    return "\n".join(out)


class MediaExtractor:
    """Copia imágenes del paquete a <media_dir>/images con nombres ordenados."""

    def __init__(self, zf: zipfile.ZipFile, media_dir: Path, md_dir: Path):
        self.zf = zf
        self.media_dir = media_dir
        self.md_dir = md_dir
        self.count = 0
        self.seen: dict = {}

    def extract(self, package_path: str) -> str | None:
        if package_path in self.seen:
            return self.seen[package_path]
        try:
            data = self.zf.read(package_path)
        except KeyError:
            return None
        self.count += 1
        ext = Path(package_path).suffix.lower() or ".png"
        dest = self.media_dir / "images" / f"figure-{self.count:03d}{ext}"
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_bytes(data)
        rel = dest.relative_to(self.md_dir)
        link = quote(str(rel))
        self.seen[package_path] = link
        return link


def shape_md(shape: ET.Element, rels: dict, media: MediaExtractor) -> tuple[str, str]:
    """(tipo, markdown) de una forma. tipo ∈ title|body|table|image|other."""
    tag = shape.tag.split("}")[1]

    if tag == "sp":
        ph = shape.find(".//p:nvSpPr/p:nvPr/p:ph", NS)
        ph_type = ph.get("type", "body") if ph is not None else "body"
        body = shape.find(".//p:txBody", NS)
        if body is None:
            return "other", ""
        lines = [paragraph_md(p, rels, as_bullets=(ph_type not in ("title", "ctrTitle")))
                 for p in body.findall("a:p", NS)]
        text = "\n".join(x for x in lines if x)
        if ph_type in ("title", "ctrTitle"):
            return "title", text.replace("\n", " ")
        return "body", text

    if tag == "graphicFrame":
        tbl = shape.find(".//a:tbl", NS)
        if tbl is not None:
            return "table", table_md(tbl, rels)
        return "other", ""

    if tag == "pic":
        blip = shape.find(".//a:blip", NS)
        if blip is not None:
            rid = blip.get(f"{{{NS['r']}}}embed")
            info = rels.get(rid)
            if info:
                link = media.extract(info["target"])
                if link:
                    return "image", f"![]({link})"
        return "other", ""

    if tag == "grpSp":  # grupo: procesar hijos recursivamente
        parts = []
        for child in shape:
            kind, md = shape_md(child, rels, media)
            if md:
                parts.append(md)
        return "body", "\n\n".join(parts)

    return "other", ""


def notes_text(zf: zipfile.ZipFile, slide_part: str, rels: dict) -> str:
    """Notas del presentador de la diapositiva (si existen)."""
    target = next((r["target"] for r in rels.values() if r["type"].endswith("/notesSlide")), None)
    if not target:
        return ""
    try:
        root = ET.fromstring(zf.read(target))
    except (KeyError, ET.ParseError):
        return ""
    lines = []
    for sp in root.iter(f"{{{NS['p']}}}sp"):
        ph = sp.find(".//p:nvSpPr/p:nvPr/p:ph", NS)
        # Ignorar placeholders de número de página y miniatura de diapositiva
        if ph is not None and ph.get("type") in ("sldNum", "sldImg"):
            continue
        for p in sp.iter(f"{{{NS['a']}}}p"):
            text = "".join(t.text or "" for t in p.iter(f"{{{NS['a']}}}t")).strip()
            if text:
                lines.append(text)
    return "\n".join(lines)


def convert(args: argparse.Namespace) -> int:
    md_path = Path(args.output)
    media_dir = Path(args.media_dir) if args.media_dir else md_path.with_name(md_path.stem + "_files")

    with zipfile.ZipFile(args.input) as zf:
        media = MediaExtractor(zf, media_dir, md_path.parent)
        sections = []

        for idx, part in enumerate(slide_parts_in_order(zf), start=1):
            rels = parse_rels(zf, part)
            try:
                root = ET.fromstring(zf.read(part))
            except (KeyError, ET.ParseError):
                continue
            tree = root.find(".//p:cSld/p:spTree", NS)
            if tree is None:
                continue

            title, blocks = "", []
            for shape in tree:
                kind, md = shape_md(shape, rels, media)
                if kind == "title" and md and not title:
                    title = md
                elif md:
                    blocks.append(md)

            lines = [f"# {title}" if title else f"# Diapositiva {idx}", ""]
            for block in blocks:
                lines += [block, ""]

            if args.notes:
                notes = notes_text(zf, part, rels)
                if notes:
                    quoted = "\n".join(f"> {ln}" for ln in notes.splitlines())
                    lines += ["> **Notas del presentador:**", ">", quoted, ""]

            sections.append("\n".join(lines).rstrip())

    if not sections:
        print("pptx_to_md: la presentación no contiene diapositivas legibles", file=sys.stderr)
        return 1

    sep = f"\n\n{args.slide_separator}\n\n" if args.slide_separator else "\n\n"
    md_path.write_text(sep.join(sections) + "\n", encoding="utf-8")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("input")
    ap.add_argument("output")
    ap.add_argument("--media-dir", default="")
    ap.add_argument("--slide-separator", default="---")
    ap.add_argument("--no-notes", dest="notes", action="store_false")
    args = ap.parse_args()
    try:
        return convert(args)
    except zipfile.BadZipFile:
        print(f"pptx_to_md: no es un paquete OOXML válido: {args.input}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())

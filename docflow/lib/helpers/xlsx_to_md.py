#!/usr/bin/env python3
"""xlsx_to_md.py — Conversor XLSX/CSV → tablas Markdown de docflow.

Pandoc no puede leer .xlsx, así que docflow parsea el contenedor OOXML
directamente (zip + XML, solo stdlib) y exporta TODAS las hojas como tablas
Markdown (una sección `##` por hoja). Con --csv procesa CSV/TSV detectando
el delimitador automáticamente.

Uso:
    xlsx_to_md.py entrada.xlsx salida.md [--max-rows 5000]
    xlsx_to_md.py entrada.csv  salida.md --csv
"""

import argparse
import csv
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from pathlib import Path

NS = {
    "s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

EXCEL_EPOCH = datetime(1899, 12, 30)


def col_index(cell_ref: str) -> int:
    """'C7' → 2 (índice de columna, base 0)."""
    letters = re.match(r"[A-Z]+", cell_ref or "A")
    idx = 0
    for ch in (letters.group() if letters else "A"):
        idx = idx * 26 + (ord(ch) - 64)
    return idx - 1


def shared_strings(zf: zipfile.ZipFile) -> list:
    try:
        root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    strings = []
    for si in root.findall("s:si", NS):
        strings.append("".join(t.text or "" for t in si.iter(f"{{{NS['s']}}}t")))
    return strings


def date_style_ids(zf: zipfile.ZipFile) -> set:
    """Índices de cellXfs cuyo numFmt es de fecha (para renderizar ISO-8601)."""
    try:
        root = ET.fromstring(zf.read("xl/styles.xml"))
    except KeyError:
        return set()
    date_fmts = {14, 15, 16, 17, 22}  # numFmtId de fecha estándar
    for fmt in root.iter(f"{{{NS['s']}}}numFmt"):
        code = (fmt.get("formatCode") or "").lower()
        if any(tok in code for tok in ("yy", "dd", "mmm")):
            date_fmts.add(int(fmt.get("numFmtId", "-1")))
    styles = set()
    xfs = root.find("s:cellXfs", NS)
    if xfs is not None:
        for i, xf in enumerate(xfs.findall("s:xf", NS)):
            if int(xf.get("numFmtId", "0")) in date_fmts:
                styles.add(i)
    return styles


def cell_value(cell: ET.Element, strings: list, date_styles: set) -> str:
    ctype = cell.get("t", "n")
    v = cell.find("s:v", NS)
    if ctype == "inlineStr":
        return "".join(t.text or "" for t in cell.iter(f"{{{NS['s']}}}t"))
    if v is None or v.text is None:
        return ""
    raw = v.text
    if ctype == "s":
        try:
            return strings[int(raw)]
        except (ValueError, IndexError):
            return raw
    if ctype == "b":
        return "VERDADERO" if raw == "1" else "FALSO"
    if ctype == "n" or ctype not in ("str", "e"):
        if int(cell.get("s", "-1")) in date_styles:
            try:
                dt = EXCEL_EPOCH + timedelta(days=float(raw))
                return dt.strftime("%Y-%m-%d") if dt.time() == dt.min.time() \
                    else dt.strftime("%Y-%m-%d %H:%M")
            except (ValueError, OverflowError):
                pass
        # Números: sin ceros decimales espurios (3.0 → 3)
        try:
            f = float(raw)
            return str(int(f)) if f.is_integer() and abs(f) < 1e15 else raw
        except ValueError:
            return raw
    return raw


def sheets_in_order(zf: zipfile.ZipFile) -> list:
    """[(nombre, ruta_parte)] en el orden del libro."""
    root = ET.fromstring(zf.read("xl/workbook.xml"))
    try:
        rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rels = {rel.get("Id"): rel.get("Target", "") for rel in rels_root.findall("rel:Relationship", NS)}
    except KeyError:
        rels = {}
    out = []
    for sheet in root.iter(f"{{{NS['s']}}}sheet"):
        rid = sheet.get(f"{{{NS['r']}}}id")
        target = rels.get(rid, "")
        if target and not target.startswith("/"):
            target = f"xl/{target}"
        out.append((sheet.get("name", "Hoja"), target.lstrip("/")))
    return out


def sheet_matrix(zf: zipfile.ZipFile, part: str, strings: list,
                 date_styles: set, max_rows: int) -> list:
    try:
        root = ET.fromstring(zf.read(part))
    except (KeyError, ET.ParseError):
        return []
    matrix, truncated = [], False
    for i, row in enumerate(root.iter(f"{{{NS['s']}}}row")):
        if i >= max_rows:
            truncated = True
            break
        cells: dict = {}
        for cell in row.findall("s:c", NS):
            cells[col_index(cell.get("r", ""))] = cell_value(cell, strings, date_styles)
        if cells:
            width = max(cells) + 1
            matrix.append([cells.get(c, "") for c in range(width)])
        else:
            matrix.append([])
    # Eliminar filas vacías al final
    while matrix and not any(x.strip() for x in matrix[-1]):
        matrix.pop()
    if truncated:
        matrix.append([f"… (truncado a {max_rows} filas)"])
    return matrix


def matrix_to_md(matrix: list) -> str:
    if not matrix:
        return "*(hoja vacía)*"
    width = max(len(r) for r in matrix)

    def fmt(row):
        cells = [c.replace("|", "\\|").replace("\n", " ") for c in row]
        cells += [""] * (width - len(cells))
        return "| " + " | ".join(cells) + " |"

    lines = [fmt(matrix[0]), "|" + "---|" * width]
    lines += [fmt(r) for r in matrix[1:]]
    return "\n".join(lines)


def convert_xlsx(args: argparse.Namespace) -> int:
    with zipfile.ZipFile(args.input) as zf:
        strings = shared_strings(zf)
        date_styles = date_style_ids(zf)
        sheets = sheets_in_order(zf)
        sections = []
        for name, part in sheets:
            if not part:
                continue
            matrix = sheet_matrix(zf, part, strings, date_styles, args.max_rows)
            sections.append(f"## {name}\n\n{matrix_to_md(matrix)}")
    if not sections:
        print(f"xlsx_to_md: libro sin hojas legibles: {args.input}", file=sys.stderr)
        return 1
    title = Path(args.input).stem
    Path(args.output).write_text(
        f"# {title}\n\n" + "\n\n".join(sections) + "\n", encoding="utf-8")
    return 0


def convert_csv(args: argparse.Namespace) -> int:
    raw = Path(args.input).read_bytes()
    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    try:
        dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel
    rows = list(csv.reader(text.splitlines(), dialect))[: args.max_rows]
    if not rows:
        print(f"xlsx_to_md: CSV vacío: {args.input}", file=sys.stderr)
        return 1
    title = Path(args.input).stem
    Path(args.output).write_text(
        f"# {title}\n\n{matrix_to_md(rows)}\n", encoding="utf-8")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("input")
    ap.add_argument("output")
    ap.add_argument("--csv", action="store_true")
    ap.add_argument("--max-rows", type=int, default=5000)
    args = ap.parse_args()
    try:
        return convert_csv(args) if args.csv else convert_xlsx(args)
    except zipfile.BadZipFile:
        print(f"xlsx_to_md: no es un paquete OOXML válido: {args.input}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())

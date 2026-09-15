#!/usr/bin/env python3
"""office_meta.py — Extractor de metadatos de documentos Office/ODF de docflow.

Lee los metadatos directamente del XML del contenedor (solo stdlib):
  · OOXML (docx/xlsx/pptx…): docProps/core.xml + docProps/app.xml
  · ODF (odt/ods/odp…):      meta.xml
Emite un bloque YAML (sin delimitadores ---) apto para frontmatter, con las
claves de meta/NORMATIVA_ARCHIVOS.md §6.2 (tipo: original; snake_case en
español). Siempre emite al menos id, tipo, titulo, estado, origen y convertido.
"""

import json
import sys
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from pathlib import Path

DC = "{http://purl.org/dc/elements/1.1/}"
DCTERMS = "{http://purl.org/dc/terms/}"
CP = "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}"
EP = "{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}"
META = "{urn:oasis:names:tc:opendocument:xmlns:meta:1.0}"
OFFICE = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"


def _text(root: ET.Element, tag: str) -> str:
    el = root.find(f".//{tag}")
    return (el.text or "").strip() if el is not None else ""


def ooxml_meta(zf: zipfile.ZipFile) -> dict:
    meta: dict = {}
    try:
        core = ET.fromstring(zf.read("docProps/core.xml"))
    except (KeyError, ET.ParseError):
        return meta
    fields = {
        "titulo": f"{DC}title",
        "autor": f"{DC}creator",
        "asunto": f"{DC}subject",
        "descripcion": f"{DC}description",
        "palabras_clave": f"{CP}keywords",
        "categoria": f"{CP}category",
        "idioma": f"{DC}language",
        "creado": f"{DCTERMS}created",
        "modificado": f"{DCTERMS}modified",
        "modificado_por": f"{CP}lastModifiedBy",
    }
    for key, tag in fields.items():
        value = _text(core, tag)
        if value:
            meta[key] = value
    try:
        app = ET.fromstring(zf.read("docProps/app.xml"))
        for key, tag in {"organizacion": f"{EP}Company", "paginas": f"{EP}Pages",
                         "palabras": f"{EP}Words", "laminas": f"{EP}Slides"}.items():
            value = _text(app, tag)
            if value and value != "0":
                meta[key] = value
    except (KeyError, ET.ParseError):
        pass
    return meta


def odf_meta(zf: zipfile.ZipFile) -> dict:
    meta: dict = {}
    try:
        root = ET.fromstring(zf.read("meta.xml"))
    except (KeyError, ET.ParseError):
        return meta
    fields = {
        "titulo": f"{DC}title",
        "autor": f"{META}initial-creator",
        "asunto": f"{DC}subject",
        "descripcion": f"{DC}description",
        "idioma": f"{DC}language",
        "creado": f"{META}creation-date",
        "modificado": f"{DC}date",
    }
    for key, tag in fields.items():
        value = _text(root, tag)
        if value:
            meta[key] = value
    keywords = [el.text.strip() for el in root.iter(f"{META}keyword") if el.text and el.text.strip()]
    if keywords:
        meta["keywords"] = ", ".join(keywords)
    return meta


def yaml_value(value: str) -> str:
    """Serialización segura: json.dumps produce escalares YAML válidos."""
    return json.dumps(value, ensure_ascii=False)


def main() -> int:
    path = Path(sys.argv[1])
    meta: dict = {}

    if zipfile.is_zipfile(path):
        with zipfile.ZipFile(path) as zf:
            names = set(zf.namelist())
            if "docProps/core.xml" in names:
                meta = ooxml_meta(zf)
            elif "meta.xml" in names:
                meta = odf_meta(zf)

    ahora = datetime.now(timezone.utc).astimezone()
    meta.setdefault("titulo", path.stem)
    meta["id"] = ahora.strftime("%Y%m%d%H%M%S")
    meta["tipo"] = "original"
    meta["estado"] = "activo"
    if meta.get("creado"):
        meta["creado"] = str(meta["creado"])[:10]          # fecha, no instante (§3)
    meta["origen"] = path.name
    meta["convertidor"] = "docflow"
    meta["convertido"] = ahora.strftime("%Y-%m-%dT%H:%M")  # instante: lo escribe una herramienta

    order = ["id", "tipo", "titulo", "estado", "creado", "autor", "asunto", "descripcion",
             "palabras_clave", "categoria", "idioma", "modificado", "modificado_por",
             "organizacion", "paginas", "palabras", "laminas", "origen", "convertidor", "convertido"]
    for key in order:
        if key in meta:
            print(f"{key}: {yaml_value(str(meta[key]))}")
    print("tags: [original]")
    print("aliases: []")
    return 0


if __name__ == "__main__":
    sys.exit(main())

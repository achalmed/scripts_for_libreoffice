#!/usr/bin/env bash
# lib/formats/doc.sh — Word legado (binario 97-2003): doc y dot.
# Pandoc no lee .doc: la estrategia via_docx lo convierte antes con LibreOffice.
registry_add doc   "document"  "Word 97-2003"                    via_docx soffice odt   ""
registry_add dot   "document"  "Plantilla Word 97-2003"          via_docx soffice ott   ""

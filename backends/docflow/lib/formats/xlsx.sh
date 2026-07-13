#!/usr/bin/env bash
# lib/formats/xlsx.sh — Excel moderno (OOXML).
# Pandoc NO lee xlsx: el parser propio (lib/helpers/xlsx_to_md.py) exporta
# todas las hojas como tablas Markdown.
registry_add xlsx  "spreadsheet" "Excel (OOXML)"                 xlsx_parser soffice ods ""
registry_add xlsm  "spreadsheet" "Excel con macros"              xlsx_parser soffice ods ""
registry_add xltx  "spreadsheet" "Plantilla Excel"               xlsx_parser soffice ots ""

#!/usr/bin/env bash
# lib/formats/ods.sh — Hojas de cálculo OpenDocument: ods, fods y ots.
registry_add ods   "spreadsheet" "Hoja de cálculo OpenDocument"  via_xlsx soffice "" xlsx
registry_add fods  "spreadsheet" "Hoja de cálculo ODF plana"     via_xlsx soffice "" xlsx
registry_add ots   "spreadsheet" "Plantilla hoja de cálculo ODF" via_xlsx soffice "" xltx

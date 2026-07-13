#!/usr/bin/env bash
# lib/formats/pdf.sh — PDF como formato de ENTRADA.
# Estrategia pdf_text: extracción con pdftotext (poppler); con --ocr aplica
# ocrmypdf antes para PDFs escaneados. Las operaciones sobre PDF (unir,
# dividir, comprimir…) viven en el PDF Engine (`docflow pdf <op>`).
registry_add pdf   "pdf" "Documento PDF"                         pdf_text none "" ""

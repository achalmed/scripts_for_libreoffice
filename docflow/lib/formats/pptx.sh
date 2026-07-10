#!/usr/bin/env bash
# lib/formats/pptx.sh — PowerPoint moderno (OOXML).
# Pandoc NO lee pptx: usamos el parser propio (lib/helpers/pptx_to_md.py)
# que extrae texto, notas del presentador, tablas e imágenes del zip OOXML.
registry_add pptx  "presentation" "PowerPoint (OOXML)"           pptx_parser soffice odp ""
registry_add pptm  "presentation" "PowerPoint con macros"        pptx_parser soffice odp ""
registry_add ppsx  "presentation" "Presentación autoejecutable"  pptx_parser soffice odp ""
registry_add potx  "presentation" "Plantilla PowerPoint"         pptx_parser soffice otp ""

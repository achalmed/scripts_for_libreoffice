#!/usr/bin/env bash
# lib/formats/odt.sh — Texto OpenDocument: odt, fodt (XML plano) y ott.
# Pandoc lee odt (zip) pero NO fodt (XML plano) → via_docx para fodt/ott.
registry_add odt   "document"  "Texto OpenDocument"              pandoc   soffice ""    docx
registry_add fodt  "document"  "Texto OpenDocument plano"        via_docx soffice ""    docx
registry_add ott   "document"  "Plantilla OpenDocument"          via_docx soffice ""    dotx

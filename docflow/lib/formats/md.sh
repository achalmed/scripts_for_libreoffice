#!/usr/bin/env bash
# lib/formats/md.sh — Markdown como formato de ENTRADA (para to-pdf/to-office).
registry_add md    "markup" "Markdown"                           none pandoc "" docx
registry_add qmd   "markup" "Quarto Markdown"                    none pandoc "" docx

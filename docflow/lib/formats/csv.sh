#!/usr/bin/env bash
# lib/formats/csv.sh — Datos tabulares planos.
registry_add csv   "data" "Valores separados por comas"          csv_parser soffice ods xlsx
registry_add tsv   "data" "Valores separados por tabuladores"    csv_parser soffice ods xlsx

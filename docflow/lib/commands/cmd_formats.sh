#!/usr/bin/env bash
# lib/commands/cmd_formats.sh — Comando formats: formatos soportados.

cmd_formats() {
    log_header "Formatos soportados"
    registry_print >&2
    log_plain ""
    log_plain "Columnas: MD = to-md · PDF = to-pdf · ODF = to-odf · MS = to-office"
    exit 0
}

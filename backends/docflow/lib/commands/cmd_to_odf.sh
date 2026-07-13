#!/usr/bin/env bash
# lib/commands/cmd_to_odf.sh — Comando to-odf: MS Office → OpenDocument.

cmd_to_odf() {
    deps_require soffice
    log_header "Conversión → OpenDocument"
    if [[ $# -eq 0 ]]; then
        log_error "Se requiere al menos una ruta."
        log_plain "  Uso: docflow to-odf [opciones] <archivo|directorio>..."
        exit "$EXIT_USAGE"
    fi
    dispatch_run odf "$@"
    exit $?
}

cli_help_to_odf() {
    cat >&2 <<EOF
${C_BOLD}docflow to-odf${C_RESET} — Convierte documentos MS Office a OpenDocument.

USO: docflow to-odf [opciones] <archivo|directorio>...

MAPA: docx/doc→odt · xlsx/xls→ods · pptx/ppt→odp
      dotx/dot→ott · xltx→ots · potx/pot→otp (plantillas)

Formatos de entrada: $(registry_exts_for odf | paste -sd' ')

EJEMPLOS:
  docflow to-odf ~/Documentos/trabajos/
  docflow to-odf -f docx,xlsx --backup ~/Documentos/legacy/
EOF
}

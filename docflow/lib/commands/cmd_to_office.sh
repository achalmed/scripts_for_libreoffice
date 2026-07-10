#!/usr/bin/env bash
# lib/commands/cmd_to_office.sh — Comando to-office: ODF → MS Office.

cmd_to_office() {
    log_header "Conversión → MS Office"
    if [[ $# -eq 0 ]]; then
        log_error "Se requiere al menos una ruta."
        log_plain "  Uso: docflow to-office [opciones] <archivo|directorio>..."
        exit "$EXIT_USAGE"
    fi
    dispatch_run office "$@"
    exit $?
}

cli_help_to_office() {
    cat >&2 <<EOF
${C_BOLD}docflow to-office${C_RESET} — Convierte OpenDocument (y Markdown) a MS Office.

USO: docflow to-office [opciones] <archivo|directorio>...

MAPA: odt→docx · ods→xlsx · odp→pptx
      ott→dotx · ots→xltx · otp→potx (plantillas)
      md→docx (vía pandoc: estilos semánticos reales)

Formatos de entrada: $(registry_exts_for ms | paste -sd' ')

EJEMPLOS:
  docflow to-office informe.odt
  docflow to-office -f odp ~/Documentos/presentaciones/
EOF
}

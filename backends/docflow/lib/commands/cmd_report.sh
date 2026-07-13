#!/usr/bin/env bash
# lib/commands/cmd_report.sh — Comando report: reportes de la última sesión.

cmd_report() {
    local formats="${DOCFLOW_REPORT_FORMATS:-html}" session
    [[ -n "${DOCFLOW_ARGS[0]:-}" ]] && formats="${DOCFLOW_ARGS[0]}"

    session="$(session_latest)"
    if [[ -z "$session" || ! -s "${session}/results.tsv" ]]; then
        log_warn "No hay sesiones de conversión registradas todavía."
        exit "$EXIT_NO_FILES"
    fi

    log_info "Sesión: $(basename -- "$session")"
    reporting_generate "$session" "$formats" "${DOCFLOW_REPORT_OUT:-}"
    exit 0
}

cli_help_report() {
    cat >&2 <<EOF
${C_BOLD}docflow report${C_RESET} — Genera reportes de la última sesión de conversión.

USO: docflow report [formatos] [--report-out <dir>]

  formatos: html, csv, json, md — separados por coma (defecto: html)

EJEMPLOS:
  docflow report
  docflow report html,json --report-out ~/reportes/
EOF
}

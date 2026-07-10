#!/usr/bin/env bash
# lib/engines/reporting.sh — Reporting Engine.
# Genera reportes HTML/CSV/JSON/Markdown de una sesión con estadísticas:
# éxito por formato, tiempos, tamaños, promedio, porcentaje. La generación
# vive en lib/helpers/report.py; aquí solo la orquestación.

# reporting_generate <session_dir> <formatos_csv> [dir_salida]
reporting_generate() {
    local session_dir="$1" formats="$2" out_dir="${3:-}"
    [[ -z "$out_dir" ]] && out_dir="${DOCFLOW_STATE_DIR}/reports"
    mkdir -p -- "$out_dir"

    if [[ ! -s "${session_dir}/results.tsv" ]]; then
        log_warn "La sesión no tiene resultados que reportar."
        return 0
    fi

    local fmt generated
    IFS=',' read -ra _fmts <<< "$formats"
    for fmt in "${_fmts[@]}"; do
        case "$fmt" in
            html|csv|json|md) ;;
            *) log_warn "Formato de reporte desconocido: $fmt"; continue ;;
        esac
        generated="$(python3 "${DOCFLOW_ROOT}/lib/helpers/report.py" \
            --session "$session_dir" --format "$fmt" --out-dir "$out_dir")" || {
            log_error "No se pudo generar el reporte ${fmt}."
            continue
        }
        log_info "Reporte ${fmt}: ${generated}"
    done
}

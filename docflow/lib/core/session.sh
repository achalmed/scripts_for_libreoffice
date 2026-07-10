#!/usr/bin/env bash
# lib/core/session.sh — Estado de una sesión de conversión.
# Cada ejecución crea un directorio de sesión con results.tsv, al que los
# workers en paralelo hacen append (líneas < PIPE_BUF ⇒ append atómico).
# Formato TSV: status  task  input  output  ms  bytes_in  bytes_out  mensaje
# status ∈ ok | skip | cached | fail | protected | dry-run

# session_init <tarea> — crea la sesión y exporta DOCFLOW_SESSION_DIR
session_init() {
    local task="$1"
    if [[ -z "${DOCFLOW_SESSION_DIR:-}" ]]; then
        DOCFLOW_SESSION_DIR="${DOCFLOW_STATE_DIR}/sessions/$(date +%Y%m%d_%H%M%S)_$$"
        mkdir -p "$DOCFLOW_SESSION_DIR"
        export DOCFLOW_SESSION_DIR
        printf '%s\n' "$task" > "${DOCFLOW_SESSION_DIR}/task"
        : > "${DOCFLOW_SESSION_DIR}/results.tsv"
        log_debug "Sesión: $DOCFLOW_SESSION_DIR"
    fi
}

# session_record <status> <task> <input> <output> <ms> <in_bytes> <out_bytes> [msg]
session_record() {
    [[ -z "${DOCFLOW_SESSION_DIR:-}" ]] && return 0
    # Los TAB de los campos de texto se sustituyen por espacios para no romper el TSV
    printf '%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\n' \
        "$1" "$2" "${3//$'\t'/ }" "${4//$'\t'/ }" "$5" "$6" "$7" "${8:-}" \
        >> "${DOCFLOW_SESSION_DIR}/results.tsv"
}

# session_count <status> → nº de resultados con ese estado ("all" = todos)
session_count() {
    local status="$1" file="${DOCFLOW_SESSION_DIR:-}/results.tsv"
    [[ -f "$file" ]] || { echo 0; return; }
    if [[ "$status" == "all" ]]; then
        wc -l < "$file"
    else
        cut -f1 "$file" | grep -c -x "$status" || true
    fi
}

# session_inputs_with_status <status> → rutas de entrada (una por línea)
session_inputs_with_status() {
    local file="${DOCFLOW_SESSION_DIR:-}/results.tsv"
    [[ -f "$file" ]] || return 0
    awk -F'\t' -v s="$1" '$1 == s { print $3 }' "$file"
}

# session_summary — resumen final; devuelve el código de salida del lote
session_summary() {
    local ok skip cached fail protected
    ok=$(session_count ok); skip=$(session_count skip)
    cached=$(session_count cached); fail=$(session_count fail)
    protected=$(session_count protected)

    log_plain ""
    log_info "Resumen → ${C_GREEN}éxito: ${ok}${C_RESET} | omitidos: ${skip} | caché: ${cached} | ${C_RED}fallos: ${fail}${C_RESET} | protegidos: ${protected}"

    if [[ "${DOCFLOW_JSON:-false}" == "true" ]]; then
        session_emit_json
    fi

    (( fail + protected > 0 )) && return "$EXIT_CONVERT_FAIL"
    return "$EXIT_OK"
}

# session_emit_json — resultados de la sesión como JSON por stdout (--json)
session_emit_json() {
    python3 "${DOCFLOW_ROOT}/lib/helpers/report.py" \
        --session "$DOCFLOW_SESSION_DIR" --format json --stdout
}

# session_latest — imprime el directorio de la sesión más reciente
session_latest() {
    local dir
    dir="$(ls -1dt "${DOCFLOW_STATE_DIR}/sessions"/*/ 2>/dev/null | head -1)"
    [[ -n "$dir" ]] && printf '%s' "${dir%/}"
}

# session_prune — conserva solo las últimas 50 sesiones
session_prune() {
    ls -1dt "${DOCFLOW_STATE_DIR}/sessions"/*/ 2>/dev/null | tail -n +51 \
        | xargs -r rm -rf
}

session_on_interrupt() {
    trap - INT TERM
    log_plain ""
    log_warn "Interrumpido. Los resultados parciales están registrados; usa --resume para continuar."
    exit "$EXIT_INTERRUPTED"
}

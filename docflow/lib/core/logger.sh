#!/usr/bin/env bash
# lib/core/logger.sh — Sistema de logging con niveles, timestamps y archivo.
# Toda la salida al usuario pasa por aquí. Los mensajes van a stderr para que
# stdout quede limpio para salida de datos (--json, pipes).
#
# Niveles: TRACE(0) DEBUG(1) INFO(2) WARN(3) ERROR(4) QUIET(5)

declare -g _LOG_LEVEL=2
declare -g _LOG_FILE=""
declare -gA _LOG_LEVEL_NAMES=([0]=TRACE [1]=DEBUG [2]=INFO [3]=WARN [4]=ERROR)

# logger_set_level <trace|debug|info|warn|error|quiet>
logger_set_level() {
    case "${1,,}" in
        trace) _LOG_LEVEL=0 ;;
        debug) _LOG_LEVEL=1 ;;
        info)  _LOG_LEVEL=2 ;;
        warn)  _LOG_LEVEL=3 ;;
        error) _LOG_LEVEL=4 ;;
        quiet) _LOG_LEVEL=5 ;;
        *) return 1 ;;
    esac
}

# logger_set_file <ruta> — activa duplicado de log a archivo (con timestamp)
logger_set_file() {
    local dir
    dir="$(dirname -- "$1")"
    mkdir -p -- "$dir" 2>/dev/null || return 1
    _LOG_FILE="$1"
    printf '# %s v%s — log iniciado %s\n' \
        "$DOCFLOW_NAME" "$DOCFLOW_VERSION" "$(date '+%Y-%m-%d %H:%M:%S')" >> "$_LOG_FILE"
}

# _log <nivel_num> <color> <etiqueta> <mensaje...>
_log() {
    local lvl="$1" color="$2" tag="$3"
    shift 3
    if [[ -n "$_LOG_FILE" ]]; then
        printf '%s [%s] %s\n' "$(date '+%Y-%m-%d %H:%M:%S')" \
            "${_LOG_LEVEL_NAMES[$lvl]}" "$*" >> "$_LOG_FILE"
    fi
    (( lvl < _LOG_LEVEL )) && return 0
    printf '%b%s%b %s\n' "$color" "$tag" "$C_RESET" "$*" >&2
}

log_trace() { _log 0 "$C_DIM"    "[trace]" "$@"; }
log_debug() { _log 1 "$C_CYAN"   "[debug]" "$@"; }
log_info()  { _log 2 "$C_BLUE"   "[info] " "$@"; }
log_warn()  { _log 3 "$C_YELLOW" "[warn] " "$@"; }
log_error() { _log 4 "$C_RED"    "[error]" "$@"; }
log_ok()    { _log 2 "$C_GREEN"  "  ✓"     "$@"; }
log_skip()  { _log 2 "$C_DIM"    "  ⊘"     "$@"; }
log_fail()  { _log 4 "$C_RED"    "  ✗"     "$@"; }

# log_plain — texto sin decorar (respeta quiet), para ayuda y listados
log_plain() {
    (( _LOG_LEVEL >= 5 )) && return 0
    printf '%s\n' "$*" >&2
}

# log_header — separador visual de secciones
log_header() {
    (( _LOG_LEVEL > 2 )) && return 0
    printf '\n%b══ %s ══%b\n' "${C_BOLD}" "$*" "${C_RESET}" >&2
    [[ -n "$_LOG_FILE" ]] && printf '== %s ==\n' "$*" >> "$_LOG_FILE"
}

#!/usr/bin/env bash
# =============================================================================
# lib/logger.sh — Sistema de logging centralizado
# Niveles: INFO · WARN · ERROR · DEBUG (solo con VERBOSE=true)
# =============================================================================
# Todas las funciones de output del proyecto pasan por aquí para garantizar
# formato uniforme y facilitar redirección a archivo si se activa --log-file.
# =============================================================================

# Guard para evitar doble-source
[[ -n "${_LOGGER_SOURCED:-}" ]] && return 0
readonly _LOGGER_SOURCED=1

# _log_write()
# Función interna que formatea y escribe una línea de log.
# No llamar directamente desde otros módulos; usar log_info/warn/error.
#
# Arguments:
#   $1 - nivel (INFO | WARN | ERROR | DEBUG)
#   $2 - color ANSI del nivel
#   $3 - mensaje
_log_write() {
    local level="$1"
    local color="$2"
    local message="$3"
    local timestamp
    timestamp="$(date '+%Y-%m-%d %H:%M:%S')"
    local formatted="${color}[${level}]${C_RESET} ${C_DIM}${timestamp}${C_RESET} — ${message}"

    # Errores y advertencias van a stderr; info y debug a stdout
    if [[ "$level" == "ERROR" || "$level" == "WARN" ]]; then
        echo -e "$formatted" >&2
    else
        echo -e "$formatted"
    fi

    # Si el logging a archivo está activo, escribe versión sin colores
    if [[ "${LOG_TO_FILE:-false}" == "true" && -n "${LOG_FILE_PATH:-}" ]]; then
        echo "[${level}] ${timestamp} — ${message}" >> "$LOG_FILE_PATH"
    fi
}

# log_info()
# Mensajes informativos del flujo normal de ejecución.
#
# Arguments:
#   $1 - mensaje a loggear
log_info() {
    _log_write "INFO " "${C_GREEN}" "$1"
}

# log_warn()
# Situaciones inesperadas que no detienen la ejecución pero merecen atención.
#
# Arguments:
#   $1 - mensaje de advertencia
log_warn() {
    _log_write "WARN " "${C_YELLOW}" "$1"
}

# log_error()
# Errores que impiden completar la operación actual.
#
# Arguments:
#   $1 - mensaje de error
log_error() {
    _log_write "ERROR" "${C_RED}" "$1"
}

# log_debug()
# Detalles técnicos útiles solo cuando VERBOSE=true.
# En producción no genera ningún output.
#
# Arguments:
#   $1 - mensaje de debug
log_debug() {
    [[ "${VERBOSE:-false}" == "true" ]] || return 0
    _log_write "DEBUG" "${C_CYAN}" "$1"
}

# log_step()
# Indica el inicio de una operación o paso visible al usuario.
# Más prominente que log_info; sirve como cabecera de cada operación.
#
# Arguments:
#   $1 - descripción del paso
log_step() {
    echo -e "\n${C_BOLD}${C_BLUE}${SYM_RUN} $1${C_RESET}"
}

# log_success()
# Resultado positivo de una operación sobre un archivo.
# Se usa después de cada procesamiento exitoso.
#
# Arguments:
#   $1 - mensaje de éxito (ej: "doc.pdf → doc_out.pdf (43% reducción)")
log_success() {
    echo -e "  ${C_GREEN}${SYM_OK}${C_RESET} $1"
}

# log_failure()
# Resultado negativo de una operación sobre un archivo.
#
# Arguments:
#   $1 - mensaje de fallo
log_failure() {
    echo -e "  ${C_RED}${SYM_FAIL}${C_RESET} $1" >&2
}

# log_skip()
# Indica que un archivo fue omitido intencionalmente (ya procesado, umbral, etc.)
#
# Arguments:
#   $1 - razón del salto
log_skip() {
    echo -e "  ${C_YELLOW}${SYM_SKIP}${C_RESET} $1"
}

# log_section()
# Imprime un separador visual con título para organizar la salida en secciones.
#
# Arguments:
#   $1 - título de la sección
log_section() {
    echo -e "\n${C_BOLD}${SEP_HEAVY}${C_RESET}"
    echo -e "${C_BOLD}  $1${C_RESET}"
    echo -e "${C_BOLD}${SEP_LIGHT}${C_RESET}"
}

# log_summary_row()
# Imprime una fila de la tabla resumen al final de una operación en lote.
#
# Arguments:
#   $1 - etiqueta   $2 - valor   $3 - color opcional
log_summary_row() {
    local label="$1"
    local value="$2"
    local color="${3:-${C_RESET}}"
    printf "  %-35s ${color}%s${C_RESET}\n" "${label}:" "$value"
}

# init_log_file()
# Crea el directorio y archivo de log si --log-file está activo.
# Debe llamarse desde main.sh después de parsear argumentos.
init_log_file() {
    [[ "${LOG_TO_FILE:-false}" == "true" ]] || return 0
    mkdir -p "$LOG_DIR" || {
        log_warn "No se pudo crear el directorio de logs: $LOG_DIR"
        LOG_TO_FILE=false
        return 1
    }
    LOG_FILE_PATH="${LOG_DIR}/pdf-suite_$(date '+%Y%m%d_%H%M%S').log"
    log_debug "Log en archivo: $LOG_FILE_PATH"
}

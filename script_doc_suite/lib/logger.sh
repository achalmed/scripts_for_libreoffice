#!/usr/bin/env bash
# lib/logger.sh — Sistema de logging centralizado
# Por qué existe: toda salida pasa por aquí para garantizar formato
# consistente, niveles de verbosidad y escritura opcional a archivo.

# Archivo de log activo (vacío = solo stdout)
LOG_FILE=""

# Acumulador de resultados para el reporte final
declare -a LOG_RESULTS=()

# ---------------------------------------------------------------------------
# _log_write()
# Función interna: formatea y escribe un mensaje. No llamar directamente.
# ---------------------------------------------------------------------------
_log_write() {
    local level="$1"
    local color="$2"
    local message="$3"
    local timestamp
    timestamp="$(date '+%Y-%m-%d %H:%M:%S')"
    local line
    printf -v line "[%-5s] %s — %s" "$level" "$timestamp" "$message"

    echo -e "${color}${line}${C_RESET}"

    # Escribir también al archivo si está configurado (sin códigos de color)
    if [[ -n "$LOG_FILE" ]]; then
        echo "$line" >> "$LOG_FILE"
    fi
}

# ---------------------------------------------------------------------------
# Funciones públicas de logging
# ---------------------------------------------------------------------------

log_info() {
    _log_write "INFO" "${C_GREEN}" "$1"
}

log_warn() {
    _log_write "WARN" "${C_YELLOW}" "$1" >&2
}

log_error() {
    _log_write "ERROR" "${C_RED}" "$1" >&2
}

# log_debug: solo imprime si VERBOSE=true
log_debug() {
    if [[ "${VERBOSE}" == "true" ]]; then
        _log_write "DEBUG" "${C_CYAN}" "$1"
    fi
}

# log_detail: línea de detalle con sangría, para listar archivos procesados
log_detail() {
    local symbol="${2:-→}"
    echo -e "  ${C_BLUE}${symbol}${C_RESET} $1"
    if [[ -n "$LOG_FILE" ]]; then
        echo "    ${symbol} $1" >> "$LOG_FILE"
    fi
}

# log_header: cabecera de sección visual
log_header() {
    local title="$1"
    local line
    printf -v line "═%.0s" {1..60}
    echo -e "\n${C_BOLD}╔${line}╗${C_RESET}"
    printf "${C_BOLD}║  %-58s║${C_RESET}\n" "$title"
    echo -e "${C_BOLD}╚${line}╝${C_RESET}"
}

# log_success: mensaje de éxito final
log_success() {
    echo -e "\n${C_GREEN}${C_BOLD}✓ $1${C_RESET}"
    if [[ -n "$LOG_FILE" ]]; then
        echo "✓ $1" >> "$LOG_FILE"
    fi
}

# ---------------------------------------------------------------------------
# record_result()
# Registra el resultado de una conversión para el reporte final.
# Argumentos: $1=archivo_origen $2=archivo_destino $3=status (ok|skip|fail) $4=modo
# ---------------------------------------------------------------------------
record_result() {
    local source="$1"
    local destination="$2"
    local status="$3"
    local mode="${4:-unknown}"
    local timestamp
    timestamp="$(date '+%Y-%m-%d %H:%M:%S')"
    # Almacenar como cadena delimitada por tabulador para parsear después
    LOG_RESULTS+=("${timestamp}\t${mode}\t${status}\t${source}\t${destination}")
}

# ---------------------------------------------------------------------------
# setup_log_file()
# Inicializa el archivo de log con una cabecera de sesión.
# Argumentos: $1=ruta del archivo de log
# ---------------------------------------------------------------------------
setup_log_file() {
    local path="$1"
    LOG_FILE="$path"
    mkdir -p "$(dirname "$LOG_FILE")"
    {
        echo "================================================================"
        echo " docflow — Log de sesión"
        echo " Inicio: $(date '+%Y-%m-%d %H:%M:%S')"
        echo " Versión: ${VERSION:-2.0.0}"
        echo "================================================================"
        echo ""
    } >> "$LOG_FILE"
    log_debug "Log activo: $LOG_FILE"
}

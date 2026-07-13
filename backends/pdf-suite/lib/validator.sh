#!/usr/bin/env bash
# =============================================================================
# lib/validator.sh — Validación de dependencias, archivos y argumentos
# =============================================================================
# Toda validación ocurre ANTES de ejecutar lógica de negocio.
# Principio: falla rápido y con mensajes útiles.
# =============================================================================

[[ -n "${_VALIDATOR_SOURCED:-}" ]] && return 0
readonly _VALIDATOR_SOURCED=1

# -----------------------------------------------------------------------------
# MAPA DE DEPENDENCIAS
# Clave: nombre del comando; valor: paquete a instalar (apt / pacman)
# Se usa para dar instrucciones precisas cuando falta una herramienta.
# -----------------------------------------------------------------------------
declare -A DEP_PACKAGES=(
    [gs]="ghostscript"
    [qpdf]="qpdf"
    [pdftk]="pdftk"
    [pdfinfo]="poppler-utils"
    [pdftotext]="poppler-utils"
    [pdfimages]="poppler-utils"
    [pdftoppm]="poppler-utils"
    [pdfunite]="poppler-utils"
    [pdfseparate]="poppler-utils"
    [pdftocairo]="poppler-utils"
    [mutool]="mupdf-tools"
    [ocrmypdf]="ocrmypdf"
    [tesseract]="tesseract-ocr"
    [exiftool]="libimage-exiftool-perl"
    [img2pdf]="img2pdf"
    [pdfjam]="texlive-extra-utils"
    [pdfcrop]="texlive-extra-utils"
    [convert]="imagemagick"
)

# -----------------------------------------------------------------------------
# DEPENDENCIAS OBLIGATORIAS Y OPCIONALES
# -----------------------------------------------------------------------------

# Herramientas sin las cuales el suite no puede funcionar en absoluto
readonly DEPS_REQUIRED=(gs qpdf pdfinfo)

# Herramientas necesarias solo para operaciones específicas
# (se validan bajo demanda en cada módulo)
readonly DEPS_OPTIONAL=(pdftk pdftotext pdfimages pdftoppm pdfunite \
    pdfseparate pdftocairo mutool ocrmypdf tesseract exiftool \
    img2pdf pdfjam pdfcrop convert)

# check_required_deps()
# Verifica que las dependencias obligatorias estén instaladas.
# Termina el programa con código 5 si alguna falta.
check_required_deps() {
    local missing=()
    for cmd in "${DEPS_REQUIRED[@]}"; do
        if ! command -v "$cmd" &>/dev/null; then
            missing+=("$cmd")
        fi
    done

    if [[ ${#missing[@]} -gt 0 ]]; then
        log_error "Faltan dependencias obligatorias:"
        for cmd in "${missing[@]}"; do
            local pkg="${DEP_PACKAGES[$cmd]:-$cmd}"
            echo -e "    ${C_RED}${SYM_FAIL}${C_RESET} ${cmd}  →  instalar: ${C_CYAN}sudo apt install ${pkg}${C_RESET}"
        done
        exit 5
    fi
    log_debug "Dependencias obligatorias: OK"
}

# check_optional_dep()
# Verifica una dependencia opcional para una operación específica.
# Retorna 1 sin matar el proceso; el módulo decide si abortar o no.
#
# Arguments:
#   $1 - nombre del comando a verificar
#   $2 - nombre de la operación que lo necesita (para el mensaje de error)
#
# Returns:
#   0 si el comando está disponible
#   1 si no está instalado (con mensaje de ayuda)
check_optional_dep() {
    local cmd="$1"
    local operation="${2:-esta operación}"

    if ! command -v "$cmd" &>/dev/null; then
        local pkg="${DEP_PACKAGES[$cmd]:-$cmd}"
        log_error "'${cmd}' es necesario para ${operation}."
        log_error "Instalar: ${C_CYAN}sudo apt install ${pkg}${C_RESET}"
        return 1
    fi
    return 0
}

# report_optional_deps()
# Muestra el estado de todas las dependencias opcionales.
# Útil para el comando --check-deps del menú principal.
report_optional_deps() {
    log_section "Estado de dependencias"
    echo ""
    printf "  %-12s %-18s %s\n" "Comando" "Paquete" "Estado"
    echo "  ${SEP_LIGHT:0:55}"

    for cmd in "${DEPS_REQUIRED[@]}"; do
        local pkg="${DEP_PACKAGES[$cmd]:-$cmd}"
        if command -v "$cmd" &>/dev/null; then
            printf "  ${C_GREEN}${SYM_OK}${C_RESET} %-10s %-18s ${C_GREEN}instalado${C_RESET}\n" "$cmd" "($pkg)"
        else
            printf "  ${C_RED}${SYM_FAIL}${C_RESET} %-10s %-18s ${C_RED}FALTA (obligatorio)${C_RESET}\n" "$cmd" "($pkg)"
        fi
    done

    echo ""
    for cmd in "${DEPS_OPTIONAL[@]}"; do
        local pkg="${DEP_PACKAGES[$cmd]:-$cmd}"
        if command -v "$cmd" &>/dev/null; then
            printf "  ${C_GREEN}${SYM_OK}${C_RESET} %-10s %-18s ${C_DIM}instalado${C_RESET}\n" "$cmd" "($pkg)"
        else
            printf "  ${C_YELLOW}${SYM_WARN}${C_RESET} %-10s %-18s ${C_YELLOW}no instalado (opcional)${C_RESET}\n" "$cmd" "($pkg)"
        fi
    done
    echo ""
}

# validate_pdf_file()
# Verifica que un path sea un archivo PDF legible y válido estructuralmente.
#
# Arguments:
#   $1 - path al archivo PDF
#
# Returns:
#   0 si es válido
#   1 si no es válido (con mensaje de error)
validate_pdf_file() {
    local file="$1"

    if [[ -z "$file" ]]; then
        log_error "No se especificó ningún archivo."
        return 1
    fi

    if [[ ! -e "$file" ]]; then
        log_error "El archivo no existe: '${file}'"
        return 1
    fi

    if [[ ! -f "$file" ]]; then
        log_error "La ruta no es un archivo regular: '${file}'"
        return 1
    fi

    if [[ ! -r "$file" ]]; then
        log_error "Sin permiso de lectura: '${file}'"
        return 1
    fi

    # Verifica magic bytes PDF (%PDF- en los primeros 5 bytes)
    local magic
    magic="$(head -c 5 "$file" 2>/dev/null)"
    if [[ "$magic" != "%PDF-" ]]; then
        log_error "El archivo no parece ser un PDF válido: '${file}'"
        return 1
    fi

    # Verifica integridad estructural básica con qpdf
    if ! qpdf --check "$file" &>/dev/null; then
        log_warn "El PDF puede tener problemas estructurales: '${file}'"
        log_warn "Continuando de todos modos; usa la opción 'reparar' si hay errores."
    fi

    return 0
}

# validate_output_path()
# Verifica que el directorio de destino exista o pueda crearse,
# y que no se sobreescriba un archivo existente salvo con --force.
#
# Arguments:
#   $1 - path completo del archivo de salida
#
# Returns:
#   0 si el path es usable
#   1 si hay un conflicto sin resolver
validate_output_path() {
    local output_path="$1"
    local output_dir
    output_dir="$(dirname "$output_path")"

    # Crear directorio si no existe
    if [[ ! -d "$output_dir" ]]; then
        if [[ "${DRY_RUN:-false}" == "true" ]]; then
            log_debug "[dry-run] Crearía directorio: $output_dir"
        else
            mkdir -p "$output_dir" || {
                log_error "No se pudo crear el directorio: $output_dir"
                return 1
            }
        fi
    fi

    # Verificar sobreescritura
    if [[ -f "$output_path" && "${FORCE:-false}" != "true" ]]; then
        log_skip "El archivo de salida ya existe (usa --force para sobreescribir): $(basename "$output_path")"
        return 1
    fi

    return 0
}

# validate_page_range()
# Verifica que un rango de páginas sea válido dado el total de páginas del PDF.
#
# Arguments:
#   $1 - rango en formato "N", "N-M", o "N,M,P"
#   $2 - total de páginas del PDF
#
# Returns:
#   0 si el rango es válido
#   1 si hay páginas fuera de rango
validate_page_range() {
    local range="$1"
    local total="$2"

    # Extrae todos los números del rango y verifica que estén dentro del total
    local invalid=false
    while IFS= read -r num; do
        if [[ "$num" =~ ^[0-9]+$ ]] && (( num < 1 || num > total )); then
            log_error "Página ${num} fuera de rango (el PDF tiene ${total} páginas)."
            invalid=true
        fi
    done < <(echo "$range" | tr ',-' '\n' | grep -E '^[0-9]+$')

    [[ "$invalid" == "true" ]] && return 1
    return 0
}

# get_page_count()
# Obtiene el número de páginas de un PDF de forma robusta.
#
# Arguments:
#   $1 - path al archivo PDF
#
# Outputs: (stdout) número de páginas, o 0 si no se puede determinar
get_page_count() {
    local file="$1"
    local count
    count="$(qpdf --show-npages "$file" 2>/dev/null)" && echo "$count" || echo "0"
}

# get_file_size_bytes()
# Obtiene el tamaño de un archivo en bytes de forma portable.
#
# Arguments:
#   $1 - path al archivo
#
# Outputs: (stdout) tamaño en bytes
get_file_size_bytes() {
    stat -c%s "$1" 2>/dev/null || echo 0
}

# format_file_size()
# Convierte bytes a formato legible (KB, MB, GB).
#
# Arguments:
#   $1 - tamaño en bytes
#
# Outputs: (stdout) tamaño formateado
format_file_size() {
    local size="$1"
    if (( size < 1024 )); then
        echo "${size} B"
    elif (( size < 1048576 )); then
        awk "BEGIN {printf \"%.1f KB\", ${size}/1024}"
    elif (( size < 1073741824 )); then
        awk "BEGIN {printf \"%.1f MB\", ${size}/1048576}"
    else
        awk "BEGIN {printf \"%.2f GB\", ${size}/1073741824}"
    fi
}

# calculate_reduction_pct()
# Calcula el porcentaje de reducción entre tamaño original y comprimido.
#
# Arguments:
#   $1 - tamaño original en bytes
#   $2 - tamaño resultante en bytes
#
# Outputs: (stdout) porcentaje de reducción (entero, puede ser negativo)
calculate_reduction_pct() {
    local original="$1"
    local result="$2"
    if (( original == 0 )); then
        echo "0"
        return
    fi
    awk "BEGIN {printf \"%d\", 100 - (${result} * 100 / ${original})}"
}

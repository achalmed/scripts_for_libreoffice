#!/usr/bin/env bash
# =============================================================================
# lib/compress.sh — Compresión de PDFs
# Motores: Ghostscript (screen/ebook/printer/prepress) + ocrmypdf (ocr)
# =============================================================================
# BUG CORREGIDO #1: stderr de gs/ocrmypdf ya no se descarta globalmente;
#   se captura en variable y se logea en modo verbose para diagnóstico.
# BUG CORREGIDO #2: contadores del lote se acumulan con archivos temporales
#   para evitar pérdida de estado en subshells creados por tuberías.
# =============================================================================

[[ -n "${_COMPRESS_SOURCED:-}" ]] && return 0
readonly _COMPRESS_SOURCED=1

# _compress_with_ghostscript()
# Comprime un PDF usando Ghostscript con el perfil de calidad indicado.
#
# Arguments:
#   $1 - archivo de entrada  $2 - archivo de salida  $3 - perfil (/ebook, etc.)
#
# Returns:
#   0 si ghostscript terminó sin error y el output es válido
#   1 si falló
_compress_with_ghostscript() {
    local input="$1"
    local output="$2"
    local profile="$3"

    log_debug "Ghostscript: perfil=${profile}, entrada=$(basename "$input")"

    local gs_stderr
    # Capturamos stderr en variable para mostrarlo solo en modo verbose,
    # en lugar de descartarlo ciegamente con 2>/dev/null.
    gs_stderr="$(gs \
        -sDEVICE=pdfwrite \
        -dCompatibilityLevel=1.4 \
        -dPDFSETTINGS="${profile}" \
        -dNOPAUSE \
        -dBATCH \
        -dQUIET \
        -sOutputFile="${output}" \
        "${input}" 2>&1 >/dev/null)"

    local exit_code=$?

    if [[ $exit_code -ne 0 ]]; then
        log_debug "Ghostscript stderr: ${gs_stderr}"
        return 1
    fi

    # Verificar que el output es un PDF válido y no vacío
    if [[ ! -s "$output" ]]; then
        log_debug "Ghostscript produjo un archivo vacío"
        return 1
    fi

    return 0
}

# _compress_with_ocrmypdf()
# Comprime un PDF usando ocrmypdf --optimize 3 con texto ya existente skipeado.
#
# Arguments:
#   $1 - archivo de entrada  $2 - archivo de salida
#
# Returns:
#   0 si ocrmypdf terminó sin error
#   1 si falló
_compress_with_ocrmypdf() {
    local input="$1"
    local output="$2"

    check_optional_dep "ocrmypdf" "compresión con método OCR" || return 1

    log_debug "ocrmypdf: optimización nivel 3, entrada=$(basename "$input")"

    local ocr_stderr
    ocr_stderr="$(ocrmypdf \
        --optimize 3 \
        --output-type pdf \
        --skip-text \
        --tesseract-timeout=0 \
        --quiet \
        "${input}" "${output}" 2>&1)"

    local exit_code=$?

    if [[ $exit_code -ne 0 ]]; then
        log_debug "ocrmypdf stderr: ${ocr_stderr}"
        return 1
    fi

    [[ -s "$output" ]] || return 1
    return 0
}

# compress_single_pdf()
# Comprime un único PDF y lo guarda como <nombre><sufijo>.pdf junto al original.
# Si la compresión aumenta el tamaño o no supera el umbral, descarta el output.
#
# Arguments:
#   $1 - path al PDF de entrada
#   $2 - método (screen|ebook|printer|prepress|ocr)
#   $3 - umbral mínimo de reducción en %
#   $4 - (opcional) path de salida explícito; si vacío, usa sufijo junto al original
#
# Returns:
#   0 éxito   1 fallo   2 saltado
compress_single_pdf() {
    local input="$1"
    local method="${2:-${COMPRESS_METHOD}}"
    local threshold="${3:-${COMPRESS_THRESHOLD}}"
    local explicit_output="${4:-}"

    # Saltar archivos que ya tienen el sufijo de salida
    if [[ "$(basename "$input")" == *"${DEFAULT_SUFFIX}.pdf" ]]; then
        log_skip "Archivo ya procesado: $(basename "$input")"
        return 2
    fi

    validate_pdf_file "$input" || return 1

    local dir
    dir="$(dirname "$input")"
    local base
    base="$(basename "$input" .pdf)"
    local output="${explicit_output:-${dir}/${base}${DEFAULT_SUFFIX}.pdf}"

    validate_output_path "$output" || return 2

    local original_size
    original_size="$(get_file_size_bytes "$input")"
    local original_fmt
    original_fmt="$(format_file_size "$original_size")"

    if [[ "${DRY_RUN:-false}" == "true" ]]; then
        log_skip "[dry-run] Comprimiría: $(basename "$input") (${original_fmt}) con método '${method}'"
        return 2
    fi

    # Archivo temporal en /tmp con nombre único para evitar colisiones
    local temp_output
    temp_output="$(mktemp /tmp/pdfsuite_compress_XXXXXX.pdf)"

    # Seleccionar motor de compresión según método
    local compress_ok=false
    case "$method" in
        screen)   _compress_with_ghostscript "$input" "$temp_output" "/screen"   && compress_ok=true ;;
        ebook)    _compress_with_ghostscript "$input" "$temp_output" "/ebook"    && compress_ok=true ;;
        printer)  _compress_with_ghostscript "$input" "$temp_output" "/printer"  && compress_ok=true ;;
        prepress) _compress_with_ghostscript "$input" "$temp_output" "/prepress" && compress_ok=true ;;
        ocr)      _compress_with_ocrmypdf    "$input" "$temp_output"             && compress_ok=true ;;
        *)
            log_error "Método de compresión no válido: '${method}'"
            rm -f "$temp_output"
            return 1
            ;;
    esac

    if [[ "$compress_ok" == "false" ]]; then
        rm -f "$temp_output"
        log_failure "Error al comprimir: $(basename "$input")"
        return 1
    fi

    local compressed_size
    compressed_size="$(get_file_size_bytes "$temp_output")"
    local compressed_fmt
    compressed_fmt="$(format_file_size "$compressed_size")"
    local reduction
    reduction="$(calculate_reduction_pct "$original_size" "$compressed_size")"

    # Descartar si la reducción no supera el umbral configurado
    if (( reduction < threshold )); then
        rm -f "$temp_output"
        log_skip "Reducción insuficiente (${reduction}% < ${threshold}%): $(basename "$input")"
        return 2
    fi

    # Descartar si el archivo resultante es más grande (regresión)
    if (( compressed_size >= original_size )); then
        rm -f "$temp_output"
        log_skip "El método '${method}' aumentó el tamaño para: $(basename "$input")"
        return 2
    fi

    mv "$temp_output" "$output"
    log_success "$(basename "$input") → $(basename "$output")  ${original_fmt} → ${compressed_fmt}  ${C_GREEN}(${reduction}% reducción)${C_RESET}"

    # Acumular estadísticas en archivo temporal para evitar pérdida en subshells
    # (corrección del BUG #2: no usar variables globales dentro de while+pipe)
    echo "$original_size $compressed_size" >> "${STATS_FILE:-/tmp/pdfsuite_stats_$$}"

    return 0
}

# compress_batch()
# Comprime todos los PDFs en un directorio (opcionalmente recursivo).
# Muestra un resumen al final con totales acumulados.
#
# Arguments:
#   $1 - directorio a procesar
#   $2 - método de compresión
#   $3 - umbral mínimo de reducción
compress_batch() {
    local target_dir="$1"
    local method="${2:-${COMPRESS_METHOD}}"
    local threshold="${3:-${COMPRESS_THRESHOLD}}"

    if [[ ! -d "$target_dir" ]]; then
        log_error "No es un directorio válido: '${target_dir}'"
        return 1
    fi

    # Archivo temporal para acumular estadísticas de archivos procesados en
    # subshells; se lee al final para construir el resumen total.
    STATS_FILE="$(mktemp /tmp/pdfsuite_stats_XXXXXX)"
    export STATS_FILE

    local count_ok=0 count_skip=0 count_fail=0 count_total=0

    log_section "Compresión en lote — método: ${method}, umbral: ${threshold}%"
    echo -e "  Directorio: ${C_CYAN}${target_dir}${C_RESET}\n"

    # Construcción del comando find según modo recursivo
    local find_args=(-type f -name "*.pdf")
    [[ "${RECURSIVE:-false}" != "true" ]] && find_args+=(-maxdepth 1)

    # Process substitution evita subshell en while, preservando variables
    while IFS= read -r pdf_file; do
        (( count_total++ ))
        echo -e "  ${C_DIM}[${count_total}]${C_RESET} $(basename "$pdf_file")"
        local result
        compress_single_pdf "$pdf_file" "$method" "$threshold"
        result=$?
        case $result in
            0) (( count_ok++ ))   ;;
            2) (( count_skip++ )) ;;
            *) (( count_fail++ )) ;;
        esac
    done < <(find "$target_dir" "${find_args[@]}" ! -name "*${DEFAULT_SUFFIX}.pdf" -print)

    # Leer estadísticas acumuladas
    local total_original=0 total_compressed=0
    if [[ -f "${STATS_FILE}" ]]; then
        while read -r orig comp; do
            (( total_original  += orig ))
            (( total_compressed += comp ))
        done < "${STATS_FILE}"
        rm -f "${STATS_FILE}"
    fi

    # Mostrar resumen
    local total_reduction=0
    (( total_original > 0 )) && \
        total_reduction="$(calculate_reduction_pct "$total_original" "$total_compressed")"

    log_section "Resumen"
    log_summary_row "Archivos procesados"          "$count_total"
    log_summary_row "Comprimidos exitosamente"      "$count_ok"     "${C_GREEN}"
    log_summary_row "Saltados"                      "$count_skip"   "${C_YELLOW}"
    log_summary_row "Fallidos"                      "$count_fail"   "${C_RED}"
    if (( total_original > 0 )); then
        echo -e "  ${SEP_LIGHT:0:55}"
        log_summary_row "Tamaño original total"    "$(format_file_size "$total_original")"
        log_summary_row "Tamaño comprimido total"  "$(format_file_size "$total_compressed")"
        log_summary_row "Espacio ahorrado"         "$(format_file_size $((total_original - total_compressed))) (${total_reduction}%)" "${C_GREEN}"
    fi
    echo ""
}

# test_all_methods()
# Prueba todos los métodos de compresión disponibles en un PDF y muestra
# una tabla comparativa. Útil para elegir el mejor método para un archivo.
#
# Arguments:
#   $1 - path al PDF de entrada
test_all_methods() {
    local input="$1"

    validate_pdf_file "$input" || return 1

    local original_size
    original_size="$(get_file_size_bytes "$input")"
    local original_fmt
    original_fmt="$(format_file_size "$original_size")"

    log_section "Test de métodos de compresión"
    echo -e "  Archivo: ${C_CYAN}$(basename "$input")${C_RESET}"
    echo -e "  Tamaño original: ${C_BOLD}${original_fmt}${C_RESET}\n"

    # Directorio temporal para los outputs de prueba; limpiado al terminar
    local test_dir
    test_dir="$(mktemp -d /tmp/pdfsuite_test_XXXXXX)"
    trap 'rm -rf "$test_dir"' RETURN

    printf "  %-12s %-14s %-12s %-22s %s\n" \
        "Método" "Tamaño" "Reducción" "Recomendación" "Comando"
    echo "  ${SEP_LIGHT}"

    local methods=(screen ebook printer prepress)
    command -v ocrmypdf &>/dev/null && methods+=(ocr)

    for method in "${methods[@]}"; do
        local out="${test_dir}/${method}.pdf"
        local ok=false

        case "$method" in
            screen)   _compress_with_ghostscript "$input" "$out" "/screen"   && ok=true ;;
            ebook)    _compress_with_ghostscript "$input" "$out" "/ebook"    && ok=true ;;
            printer)  _compress_with_ghostscript "$input" "$out" "/printer"  && ok=true ;;
            prepress) _compress_with_ghostscript "$input" "$out" "/prepress" && ok=true ;;
            ocr)      _compress_with_ocrmypdf    "$input" "$out"             && ok=true ;;
        esac

        if [[ "$ok" == "true" && -s "$out" ]]; then
            local size fmt reduction
            size="$(get_file_size_bytes "$out")"
            fmt="$(format_file_size "$size")"
            reduction="$(calculate_reduction_pct "$original_size" "$size")"

            local recommendation=""
            case "$method" in
                screen)   recommendation="Web / email" ;;
                ebook)    recommendation="Lectura digital ⭐" ;;
                printer)  recommendation="Imprimir" ;;
                prepress) recommendation="Impresión pro" ;;
                ocr)      recommendation="Escaneados ⭐" ;;
            esac

            local color="${C_RESET}"
            (( reduction >= 40 )) && color="${C_GREEN}"
            (( reduction < 0  )) && color="${C_RED}"

            printf "  %-12s %-14s ${color}%9s%%%${C_RESET}  %-22s %s\n" \
                "$method" "$fmt" "$reduction" "$recommendation" \
                "pdf-suite compress -m ${method}"
        else
            printf "  %-12s ${C_RED}%-14s${C_RESET}\n" "$method" "FALLÓ"
        fi
    done

    echo ""
    echo -e "  ${C_DIM}Para aplicar: pdf-suite compress -m <método> \"$(basename "$input")\"${C_RESET}"
    echo ""
}

# run_compress()
# Punto de entrada del módulo de compresión; parsea los argumentos específicos
# de la operación y delega a compress_single_pdf o compress_batch.
#
# Arguments:
#   "$@" - argumentos restantes después de extraer flags globales
run_compress() {
    local method="$COMPRESS_METHOD"
    local threshold="$COMPRESS_THRESHOLD"
    local targets=()

    while [[ $# -gt 0 ]]; do
        case "$1" in
            -m|--method)    method="$2";    shift 2 ;;
            -t|--threshold) threshold="$2"; shift 2 ;;
            -h|--help)
                echo "Uso: pdf-suite compress [-m MÉTODO] [-t UMBRAL] <archivo|directorio>"
                echo "Métodos: screen ebook printer prepress ocr"
                return 0 ;;
            *) targets+=("$1"); shift ;;
        esac
    done

    if [[ ${#targets[@]} -eq 0 ]]; then
        # Modo interactivo
        prompt_for_file "Ruta del PDF o carpeta a comprimir"
        targets=("$PROMPTED_FILE")
        prompt_for_option "Método (screen|ebook|printer|prepress|ocr)" "$method"
        method="$PROMPTED_VALUE"
        prompt_for_option "Umbral mínimo de reducción (%)" "$threshold"
        threshold="$PROMPTED_VALUE"
    fi

    check_optional_dep "gs" "compresión" || return 1

    for target in "${targets[@]}"; do
        if [[ -f "$target" ]]; then
            log_step "Comprimiendo: $(basename "$target")"
            compress_single_pdf "$target" "$method" "$threshold" "${OUTPUT_PATH:-}"
        elif [[ -d "$target" ]]; then
            compress_batch "$target" "$method" "$threshold"
        else
            log_error "Objetivo no encontrado: '${target}'"
            return 1
        fi
    done
}

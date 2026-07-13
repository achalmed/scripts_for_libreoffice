#!/usr/bin/env bash
# =============================================================================
# lib/ocr.sh — OCR para PDFs escaneados
# Motor principal : ocrmypdf (añade capa de texto buscable sobre el PDF)
# Motor secundario: tesseract (para imágenes directas)
# =============================================================================

[[ -n "${_OCR_SOURCED:-}" ]] && return 0
readonly _OCR_SOURCED=1

# apply_ocr()
# Aplica OCR a un PDF escaneado añadiendo una capa de texto buscable.
# Preserva la imagen original; no recomprime.
#
# Arguments:
#   $1 - archivo PDF de entrada
#   $2 - idioma(s) Tesseract (default: spa; múltiples: "spa+eng")
#   $3 - (opcional) path de salida
apply_ocr() {
    local input="$1"
    local lang="${2:-${OCR_LANG}}"
    local output="${3:-}"

    check_optional_dep "ocrmypdf" "OCR de PDFs" || return 1
    validate_pdf_file "$input" || return 1

    local dir base
    dir="$(dirname "$input")"
    base="$(basename "$input" .pdf)"
    output="${output:-${dir}/${base}_ocr.pdf}"

    validate_output_path "$output" || return 1

    log_step "OCR: $(basename "$input") (idioma: ${lang})"

    if [[ "${DRY_RUN:-false}" == "true" ]]; then
        log_skip "[dry-run] Crearía: $(basename "$output")"
        return 0
    fi

    local ocr_stderr
    ocr_stderr="$(ocrmypdf \
        -l "$lang" \
        --deskew \
        --clean \
        --rotate-pages \
        --skip-text \
        --output-type pdf \
        --quiet \
        "$input" "$output" 2>&1)"

    local exit_code=$?

    if [[ $exit_code -ne 0 ]]; then
        log_debug "ocrmypdf stderr: ${ocr_stderr}"
        log_failure "Error al aplicar OCR a $(basename "$input")."
        rm -f "$output"
        return 1
    fi

    local orig_sz comp_sz reduction
    orig_sz="$(format_file_size "$(get_file_size_bytes "$input")")"
    comp_sz="$(format_file_size "$(get_file_size_bytes "$output")")"
    log_success "$(basename "$output")  ${orig_sz} → ${comp_sz}  (texto buscable añadido)"
}

# apply_ocr_batch()
# Aplica OCR a todos los PDFs de un directorio.
#
# Arguments:
#   $1 - directorio
#   $2 - idioma
apply_ocr_batch() {
    local target_dir="$1"
    local lang="${2:-${OCR_LANG}}"

    if [[ ! -d "$target_dir" ]]; then
        log_error "No es un directorio: '${target_dir}'"
        return 1
    fi

    log_section "OCR en lote — idioma: ${lang}"
    echo -e "  Directorio: ${C_CYAN}${target_dir}${C_RESET}\n"

    local count=0 ok=0 fail=0

    local find_args=(-type f -name "*.pdf")
    [[ "${RECURSIVE:-false}" != "true" ]] && find_args+=(-maxdepth 1)

    while IFS= read -r pdf; do
        [[ "$(basename "$pdf")" == *_ocr.pdf ]] && continue
        (( count++ ))
        echo -e "  ${C_DIM}[${count}]${C_RESET} $(basename "$pdf")"
        if apply_ocr "$pdf" "$lang"; then
            (( ok++ ))
        else
            (( fail++ ))
        fi
    done < <(find "$target_dir" "${find_args[@]}" -print)

    log_section "Resumen OCR"
    log_summary_row "Procesados" "$count"
    log_summary_row "Exitosos"   "$ok"   "${C_GREEN}"
    log_summary_row "Fallidos"   "$fail" "${C_RED}"
    echo ""
}

# check_if_needs_ocr()
# Detecta si un PDF tiene texto o es solo imagen (necesita OCR).
# Comprueba extrayendo texto con pdftotext y contando palabras.
#
# Arguments:
#   $1 - archivo PDF
#
# Outputs: (stdout) "yes" si necesita OCR, "no" si ya tiene texto
check_if_needs_ocr() {
    local input="$1"

    check_optional_dep "pdftotext" "detección de OCR" || { echo "unknown"; return 1; }

    local word_count
    word_count="$(pdftotext "$input" - 2>/dev/null | wc -w)"

    if (( word_count < 10 )); then
        echo "yes"
    else
        echo "no"
    fi
}

# scan_for_ocr_needed()
# Escanea un directorio e indica qué PDFs probablemente son escaneados
# y necesitan OCR (tienen menos de 10 palabras extraíbles).
#
# Arguments:
#   $1 - directorio a escanear
scan_for_ocr_needed() {
    local target_dir="${1:-.}"

    check_optional_dep "pdftotext" "escaneo OCR" || return 1

    log_section "PDFs que probablemente necesitan OCR en: $target_dir"

    local find_args=(-type f -name "*.pdf")
    [[ "${RECURSIVE:-false}" != "true" ]] && find_args+=(-maxdepth 1)

    local count_needs=0 count_ok=0

    while IFS= read -r pdf; do
        local needs
        needs="$(check_if_needs_ocr "$pdf")"
        if [[ "$needs" == "yes" ]]; then
            echo -e "  ${C_YELLOW}${SYM_WARN}${C_RESET} $(basename "$pdf")"
            (( count_needs++ ))
        else
            (( count_ok++ ))
        fi
    done < <(find "$target_dir" "${find_args[@]}" -print)

    echo ""
    log_summary_row "Ya tienen texto" "$count_ok"    "${C_GREEN}"
    log_summary_row "Necesitan OCR"   "$count_needs" "${C_YELLOW}"
    echo ""
    (( count_needs > 0 )) && \
        echo -e "  ${C_DIM}Para aplicar: pdf-suite ocr -r '${target_dir}'${C_RESET}"
    echo ""
}

# run_ocr()
# Punto de entrada del módulo OCR.
run_ocr() {
    local lang="$OCR_LANG"
    local scan_mode=false
    local targets=()

    while [[ $# -gt 0 ]]; do
        case "$1" in
            -l|--lang)    lang="$2"; shift 2 ;;
            --scan)       scan_mode=true; shift ;;
            *) targets+=("$1"); shift ;;
        esac
    done

    if [[ "$scan_mode" == "true" ]]; then
        local dir="${targets[0]:-.}"
        scan_for_ocr_needed "$dir"
        return
    fi

    if [[ ${#targets[@]} -eq 0 ]]; then
        prompt_for_file "Ruta del PDF o directorio a procesar con OCR"
        targets=("$PROMPTED_FILE")
        prompt_for_option "Idioma(s) Tesseract" "$lang"
        lang="$PROMPTED_VALUE"
    fi

    for t in "${targets[@]}"; do
        if [[ -d "$t" ]]; then
            apply_ocr_batch "$t" "$lang"
        else
            apply_ocr "$t" "$lang" "${OUTPUT_PATH:-}"
        fi
    done
}

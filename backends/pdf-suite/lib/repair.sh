#!/usr/bin/env bash
# =============================================================================
# lib/repair.sh — Reparación, validación y optimización de PDFs
# Motores: qpdf · mutool · ghostscript
# =============================================================================

[[ -n "${_REPAIR_SOURCED:-}" ]] && return 0
readonly _REPAIR_SOURCED=1

# validate_pdf_structure()
# Realiza una validación completa de la estructura interna del PDF.
# Reporta problemas encontrados sin modificar el archivo.
#
# Arguments:
#   $1 - archivo PDF
validate_pdf_structure() {
    local input="$1"

    validate_pdf_file "$input" || return 1

    log_step "Validando estructura de: $(basename "$input")"
    echo ""

    local has_errors=false

    # qpdf --check: validación estructural completa
    local qpdf_out
    qpdf_out="$(qpdf --check "$input" 2>&1)"
    local qpdf_exit=$?

    if [[ $qpdf_exit -eq 0 ]]; then
        echo -e "  ${C_GREEN}${SYM_OK}${C_RESET} qpdf: estructura válida"
    else
        echo -e "  ${C_RED}${SYM_FAIL}${C_RESET} qpdf: problemas detectados"
        echo "$qpdf_out" | grep -i "error\|warning" | sed 's/^/    /'
        has_errors=true
    fi

    # Verificar si tiene contraseña
    if qpdf --check "$input" 2>&1 | grep -qi "password"; then
        echo -e "  ${C_YELLOW}${SYM_WARN}${C_RESET} El PDF está protegido con contraseña"
    fi

    # pdfinfo: metadatos e información
    if command -v pdfinfo &>/dev/null; then
        local pdfinfo_out
        pdfinfo_out="$(pdfinfo "$input" 2>&1)"
        if echo "$pdfinfo_out" | grep -qi "error"; then
            echo -e "  ${C_YELLOW}${SYM_WARN}${C_RESET} pdfinfo reportó problemas"
        else
            local pages
            pages="$(echo "$pdfinfo_out" | grep "^Pages:" | awk '{print $2}')"
            local version
            version="$(echo "$pdfinfo_out" | grep "^PDF version:" | awk '{print $3}')"
            echo -e "  ${C_GREEN}${SYM_OK}${C_RESET} pdfinfo: ${pages} páginas, PDF v${version}"
        fi
    fi

    # mutool info: verificación adicional
    if command -v mutool &>/dev/null; then
        if mutool info "$input" &>/dev/null; then
            echo -e "  ${C_GREEN}${SYM_OK}${C_RESET} mutool: estructura reconocida"
        else
            echo -e "  ${C_YELLOW}${SYM_WARN}${C_RESET} mutool no reconoce el archivo correctamente"
        fi
    fi

    echo ""
    if [[ "$has_errors" == "true" ]]; then
        echo -e "  ${C_YELLOW}Recomendación: ejecuta 'pdf-suite repair $(basename "$input")'${C_RESET}"
    else
        echo -e "  ${C_GREEN}El PDF parece estructuralmente correcto.${C_RESET}"
    fi
    echo ""
}

# repair_pdf()
# Intenta reparar un PDF dañado usando múltiples estrategias en cascada.
# Estrategia 1 (qpdf): reescribe el PDF respetando la estructura existente.
# Estrategia 2 (mutool clean): reconstruye objetos dañados.
# Estrategia 3 (ghostscript): re-renderiza completamente (último recurso).
#
# Arguments:
#   $1 - archivo PDF de entrada
#   $2 - (opcional) path de salida
repair_pdf() {
    local input="$1"
    local output="${2:-}"

    validate_pdf_file "$input" 2>/dev/null || {
        log_warn "El archivo puede estar corrupto; intentando reparar de todos modos."
    }

    local dir base
    dir="$(dirname "$input")"
    base="$(basename "$input" .pdf)"
    output="${output:-${dir}/${base}_repaired.pdf}"

    validate_output_path "$output" || return 1

    log_step "Reparando: $(basename "$input")"

    if [[ "${DRY_RUN:-false}" == "true" ]]; then
        log_skip "[dry-run] Crearía: $(basename "$output")"
        return 0
    fi

    local temp
    temp="$(mktemp /tmp/pdfsuite_repair_XXXXXX.pdf)"

    # Estrategia 1: qpdf --linearize (reescribe y lineariza)
    log_debug "Estrategia 1: qpdf"
    if qpdf --linearize "$input" "$temp" 2>/dev/null && [[ -s "$temp" ]]; then
        mv "$temp" "$output"
        log_success "$(basename "$output") reparado con qpdf"
        return 0
    fi

    # Estrategia 2: mutool clean (reconstruye objetos)
    if command -v mutool &>/dev/null; then
        log_debug "Estrategia 2: mutool clean"
        if mutool clean -g "$input" "$temp" 2>/dev/null && [[ -s "$temp" ]]; then
            mv "$temp" "$output"
            log_success "$(basename "$output") reparado con mutool"
            return 0
        fi
    fi

    # Estrategia 3: Ghostscript (re-renderizado completo)
    log_debug "Estrategia 3: ghostscript (re-renderizado completo)"
    if gs -dBATCH -dNOPAUSE -sDEVICE=pdfwrite \
           -dSAFER \
           -sOutputFile="$temp" \
           "$input" &>/dev/null && [[ -s "$temp" ]]; then
        mv "$temp" "$output"
        log_success "$(basename "$output") reparado con Ghostscript (re-renderizado)"
        log_warn "El PDF fue re-renderizado; pueden perderse marcadores y formularios."
        return 0
    fi

    rm -f "$temp"
    log_failure "No se pudo reparar $(basename "$input"). El archivo puede estar irrecuperable."
    return 1
}

# optimize_pdf()
# Optimiza un PDF para visualización web (linearización fast web view)
# y compresión de streams internos sin cambiar calidad visual.
#
# Arguments:
#   $1 - archivo de entrada
#   $2 - (opcional) path de salida
optimize_pdf() {
    local input="$1"
    local output="${2:-}"

    validate_pdf_file "$input" || return 1

    local dir base
    dir="$(dirname "$input")"
    base="$(basename "$input" .pdf)"
    output="${output:-${dir}/${base}_optimized.pdf}"

    validate_output_path "$output" || return 1

    log_step "Optimizando estructura: $(basename "$input")"

    if [[ "${DRY_RUN:-false}" == "true" ]]; then
        log_skip "[dry-run] Crearía: $(basename "$output")"
        return 0
    fi

    # qpdf --linearize: optimiza para "Fast Web View" (carga progresiva)
    # --compress-streams: comprime streams no comprimidos
    if qpdf --linearize \
            --compress-streams=y \
            --recompress-flate \
            "$input" "$output" 2>/dev/null; then
        local orig_sz new_sz reduction
        orig_sz="$(get_file_size_bytes "$input")"
        new_sz="$(get_file_size_bytes "$output")"
        reduction="$(calculate_reduction_pct "$orig_sz" "$new_sz")"
        log_success "$(basename "$output")  (Fast Web View + streams comprimidos, ${reduction}% reducción)"
    else
        log_failure "Error al optimizar $(basename "$input")."
        rm -f "$output"
        return 1
    fi
}

# run_repair()
# Punto de entrada del módulo de reparación.
run_repair() {
    local mode="repair"
    local targets=()

    while [[ $# -gt 0 ]]; do
        case "$1" in
            --validate)  mode="validate"; shift ;;
            --optimize)  mode="optimize"; shift ;;
            *) targets+=("$1"); shift ;;
        esac
    done

    if [[ ${#targets[@]} -eq 0 ]]; then
        prompt_for_file "Ruta del PDF"
        targets=("$PROMPTED_FILE")
        prompt_for_option "Operación (repair|validate|optimize)" "repair"
        mode="$PROMPTED_VALUE"
    fi

    for t in "${targets[@]}"; do
        case "$mode" in
            repair)   repair_pdf "$t" "${OUTPUT_PATH:-}" ;;
            validate) validate_pdf_structure "$t" ;;
            optimize) optimize_pdf "$t" "${OUTPUT_PATH:-}" ;;
        esac
    done
}

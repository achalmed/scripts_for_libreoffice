#!/usr/bin/env bash
# =============================================================================
# lib/manipulate.sh — Manipulación estructural de PDFs
# Operaciones: merge · split · extract · rotate · reorder · delete
# Motor principal: qpdf (preserva contenido sin recomprimir)
# Motor secundario: pdfunite / pdfseparate (operaciones simples de poppler)
# =============================================================================

[[ -n "${_MANIPULATE_SOURCED:-}" ]] && return 0
readonly _MANIPULATE_SOURCED=1

# merge_pdfs()
# Une múltiples PDFs en un único archivo de salida.
# Preserva marcadores, formularios y metadatos del primer PDF.
#
# Arguments:
#   $1    - path de salida
#   $2..N - archivos de entrada en el orden deseado
merge_pdfs() {
    local output="$1"
    shift
    local inputs=("$@")

    if [[ ${#inputs[@]} -lt 2 ]]; then
        log_error "Se necesitan al menos 2 PDFs para unir."
        return 1
    fi

    for f in "${inputs[@]}"; do
        validate_pdf_file "$f" || return 1
    done

    validate_output_path "$output" || return 1

    log_step "Uniendo ${#inputs[@]} PDFs → $(basename "$output")"

    if [[ "${DRY_RUN:-false}" == "true" ]]; then
        log_skip "[dry-run] Uniría: ${inputs[*]}"
        return 0
    fi

    # Construir argumentos de qpdf: --empty --pages file1 file2 ... --
    local qpdf_args=(--empty --pages)
    for f in "${inputs[@]}"; do
        qpdf_args+=("$f")
    done
    qpdf_args+=(-- "$output")

    if qpdf "${qpdf_args[@]}" 2>/dev/null; then
        local out_size
        out_size="$(format_file_size "$(get_file_size_bytes "$output")")"
        log_success "$(basename "$output") creado (${out_size}, ${#inputs[@]} archivos)"
    else
        log_failure "Error al unir los PDFs."
        return 1
    fi
}

# split_pdf()
# Divide un PDF en partes de N páginas cada una.
# Los archivos de salida siguen el patrón: <nombre>_parte001.pdf, etc.
#
# Arguments:
#   $1 - archivo de entrada
#   $2 - páginas por parte (default: 1 = una página por archivo)
split_pdf() {
    local input="$1"
    local pages_per_part="${2:-1}"

    validate_pdf_file "$input" || return 1

    local total
    total="$(get_page_count "$input")"
    if (( total == 0 )); then
        log_error "No se pudo obtener el número de páginas de '$(basename "$input")'."
        return 1
    fi

    local dir base
    dir="$(dirname "$input")"
    base="$(basename "$input" .pdf)"
    local output_prefix="${dir}/${base}_parte"

    log_step "Dividiendo $(basename "$input") (${total} págs) en grupos de ${pages_per_part}"

    if [[ "${DRY_RUN:-false}" == "true" ]]; then
        local parts=$(( (total + pages_per_part - 1) / pages_per_part ))
        log_skip "[dry-run] Crearía ${parts} archivos con prefijo $(basename "$output_prefix")"
        return 0
    fi

    local part=1 start=1
    while (( start <= total )); do
        local end=$(( start + pages_per_part - 1 ))
        (( end > total )) && end=$total

        local out_file
        out_file="$(printf "%s%03d.pdf" "$output_prefix" "$part")"

        if qpdf --empty --pages "$input" "${start}-${end}" -- "$out_file" 2>/dev/null; then
            local sz
            sz="$(format_file_size "$(get_file_size_bytes "$out_file")")"
            log_success "$(basename "$out_file")  págs ${start}–${end}  (${sz})"
        else
            log_failure "Error en parte ${part} (págs ${start}–${end})"
        fi

        (( start = end + 1 ))
        (( part++ ))
    done
}

# extract_pages()
# Extrae páginas específicas de un PDF usando rangos flexibles.
# Soporta: "3-7", "1,5,9", "1-3,7,10-z" (z = última página).
#
# Arguments:
#   $1 - archivo de entrada
#   $2 - rango de páginas (ej: "1-5,10,15-z")
#   $3 - (opcional) path de salida explícito
extract_pages() {
    local input="$1"
    local page_range="$2"
    local output="${3:-}"

    if [[ -z "$page_range" ]]; then
        log_error "Debes especificar un rango de páginas (ej: --pages '1-5,10')."
        return 1
    fi

    validate_pdf_file "$input" || return 1

    local total
    total="$(get_page_count "$input")"

    # Validar el rango (solo para rangos sin 'z')
    local range_no_z="${page_range//z/$total}"
    validate_page_range "$range_no_z" "$total" || return 1

    local dir base
    dir="$(dirname "$input")"
    base="$(basename "$input" .pdf)"
    output="${output:-${dir}/${base}_pp${page_range//,/_}${DEFAULT_SUFFIX}.pdf}"
    # Limpiar caracteres no válidos en el nombre de archivo
    output="${output//[[:space:]]/_}"

    validate_output_path "$output" || return 1

    log_step "Extrayendo páginas '${page_range}' de $(basename "$input")"

    if [[ "${DRY_RUN:-false}" == "true" ]]; then
        log_skip "[dry-run] Crearía: $(basename "$output")"
        return 0
    fi

    if qpdf --empty --pages "$input" "$page_range" -- "$output" 2>/dev/null; then
        local extracted_count
        extracted_count="$(get_page_count "$output")"
        local sz
        sz="$(format_file_size "$(get_file_size_bytes "$output")")"
        log_success "$(basename "$output")  (${extracted_count} págs, ${sz})"
    else
        log_failure "Error al extraer páginas de $(basename "$input")."
        rm -f "$output"
        return 1
    fi
}

# rotate_pages()
# Rota páginas de un PDF. Soporta rotación de todas las páginas o rangos.
#
# Arguments:
#   $1 - archivo de entrada
#   $2 - ángulo: 90, 180, 270 (o +90/-90 para relativo)
#   $3 - (opcional) rango de páginas; si vacío, rota todas
#   $4 - (opcional) path de salida
rotate_pages() {
    local input="$1"
    local angle="$2"
    local page_range="${3:-}"
    local output="${4:-}"

    if [[ -z "$angle" ]]; then
        log_error "Debes especificar un ángulo (90, 180, 270)."
        return 1
    fi

    # Normalizar ángulo al formato de qpdf (+N:rango o +N)
    local qpdf_rotate
    if [[ -n "$page_range" ]]; then
        qpdf_rotate="+${angle#-}:${page_range}"
        # Si el ángulo original era negativo, usar negativo
        [[ "$angle" == -* ]] && qpdf_rotate="${angle}:${page_range}"
    else
        qpdf_rotate="+${angle#-}"
        [[ "$angle" == -* ]] && qpdf_rotate="$angle"
    fi

    validate_pdf_file "$input" || return 1

    local dir base
    dir="$(dirname "$input")"
    base="$(basename "$input" .pdf)"
    output="${output:-${dir}/${base}_rot${angle}${DEFAULT_SUFFIX}.pdf}"

    validate_output_path "$output" || return 1

    log_step "Rotando $(basename "$input") ${angle}°${page_range:+ (págs: $page_range)}"

    if [[ "${DRY_RUN:-false}" == "true" ]]; then
        log_skip "[dry-run] Crearía: $(basename "$output")"
        return 0
    fi

    if qpdf --rotate="${qpdf_rotate}" "$input" "$output" 2>/dev/null; then
        log_success "$(basename "$output") creado"
    else
        log_failure "Error al rotar $(basename "$input")."
        rm -f "$output"
        return 1
    fi
}

# reorder_pages()
# Reordena las páginas de un PDF según un rango qpdf.
# Caso de uso más común: invertir el orden con rango "z-1".
#
# Arguments:
#   $1 - archivo de entrada
#   $2 - rango qpdf de reordenamiento (ej: "z-1" para invertir)
#   $3 - (opcional) path de salida
reorder_pages() {
    local input="$1"
    local reorder_range="${2:-z-1}"
    local output="${3:-}"

    validate_pdf_file "$input" || return 1

    local dir base
    dir="$(dirname "$input")"
    base="$(basename "$input" .pdf)"
    output="${output:-${dir}/${base}_reordered${DEFAULT_SUFFIX}.pdf}"

    validate_output_path "$output" || return 1

    log_step "Reordenando páginas de $(basename "$input") (rango: ${reorder_range})"

    if [[ "${DRY_RUN:-false}" == "true" ]]; then
        log_skip "[dry-run] Crearía: $(basename "$output")"
        return 0
    fi

    if qpdf --empty --pages "$input" "$reorder_range" -- "$output" 2>/dev/null; then
        log_success "$(basename "$output") creado"
    else
        log_failure "Error al reordenar $(basename "$input")."
        rm -f "$output"
        return 1
    fi
}

# delete_pages()
# Elimina páginas de un PDF indicando cuáles quitar.
# Internamente calcula las páginas que SÍ se mantienen.
#
# Arguments:
#   $1 - archivo de entrada
#   $2 - páginas a ELIMINAR (ej: "3,5,10-12")
#   $3 - (opcional) path de salida
delete_pages() {
    local input="$1"
    local pages_to_delete="$2"
    local output="${3:-}"

    if [[ -z "$pages_to_delete" ]]; then
        log_error "Debes indicar qué páginas eliminar (ej: --pages '3,5,10-12')."
        return 1
    fi

    validate_pdf_file "$input" || return 1

    local total
    total="$(get_page_count "$input")"

    # Construir el conjunto de páginas a MANTENER
    # Expandimos las páginas a eliminar y luego las excluimos del rango total
    local delete_set
    delete_set="$(python3 -c "
import sys
total = $total
raw = '$pages_to_delete'
delete = set()
for part in raw.split(','):
    part = part.strip()
    if '-' in part:
        a, b = part.split('-')
        delete.update(range(int(a), int(b)+1))
    else:
        delete.add(int(part))
keep = [str(p) for p in range(1, total+1) if p not in delete]
print(','.join(keep))
" 2>/dev/null)"

    if [[ -z "$delete_set" ]]; then
        log_error "No quedarían páginas después de eliminar '${pages_to_delete}'."
        return 1
    fi

    local dir base
    dir="$(dirname "$input")"
    base="$(basename "$input" .pdf)"
    output="${output:-${dir}/${base}_deleted${DEFAULT_SUFFIX}.pdf}"

    validate_output_path "$output" || return 1

    log_step "Eliminando páginas '${pages_to_delete}' de $(basename "$input") (${total} págs)"

    if [[ "${DRY_RUN:-false}" == "true" ]]; then
        log_skip "[dry-run] Crearía: $(basename "$output")"
        return 0
    fi

    if qpdf --empty --pages "$input" "$delete_set" -- "$output" 2>/dev/null; then
        local remaining
        remaining="$(get_page_count "$output")"
        log_success "$(basename "$output")  (${remaining} págs restantes)"
    else
        log_failure "Error al eliminar páginas de $(basename "$input")."
        rm -f "$output"
        return 1
    fi
}

# ---------------------------------------------------------------------------
# Punto de entrada para cada sub-operación del módulo
# ---------------------------------------------------------------------------

run_merge() {
    local output="${OUTPUT_PATH:-}"
    local inputs=()
    while [[ $# -gt 0 ]]; do
        case "$1" in
            -o|--output) output="$2"; shift 2 ;;
            *) inputs+=("$1"); shift ;;
        esac
    done
    if [[ ${#inputs[@]} -eq 0 ]]; then
        echo "Uso: pdf-suite merge [-o salida.pdf] archivo1.pdf archivo2.pdf ..."
        return 1
    fi
    [[ -z "$output" ]] && output="$(dirname "${inputs[0]}")/merged${DEFAULT_SUFFIX}.pdf"
    merge_pdfs "$output" "${inputs[@]}"
}

run_split() {
    local pages_per_part=1
    local targets=()
    while [[ $# -gt 0 ]]; do
        case "$1" in
            --pages|-p) pages_per_part="$2"; shift 2 ;;
            *) targets+=("$1"); shift ;;
        esac
    done
    if [[ ${#targets[@]} -eq 0 ]]; then
        prompt_for_file "Ruta del PDF a dividir"
        targets=("$PROMPTED_FILE")
        prompt_for_option "Páginas por parte" "1"
        pages_per_part="$PROMPTED_VALUE"
    fi
    for t in "${targets[@]}"; do
        split_pdf "$t" "$pages_per_part"
    done
}

run_extract() {
    local page_range=""
    local targets=()
    while [[ $# -gt 0 ]]; do
        case "$1" in
            --pages|-p) page_range="$2"; shift 2 ;;
            *) targets+=("$1"); shift ;;
        esac
    done
    if [[ ${#targets[@]} -eq 0 || -z "$page_range" ]]; then
        prompt_for_file "Ruta del PDF"
        targets=("$PROMPTED_FILE")
        prompt_for_option "Rango de páginas (ej: 1-5,10,15-z)" "1-z"
        page_range="$PROMPTED_VALUE"
    fi
    for t in "${targets[@]}"; do
        extract_pages "$t" "$page_range" "${OUTPUT_PATH:-}"
    done
}

run_rotate() {
    local angle="" page_range="" targets=()
    while [[ $# -gt 0 ]]; do
        case "$1" in
            --angle|-a) angle="$2"; shift 2 ;;
            --pages|-p) page_range="$2"; shift 2 ;;
            *) targets+=("$1"); shift ;;
        esac
    done
    if [[ ${#targets[@]} -eq 0 || -z "$angle" ]]; then
        prompt_for_file "Ruta del PDF"
        targets=("$PROMPTED_FILE")
        prompt_for_option "Ángulo (90, 180, 270)" "90"
        angle="$PROMPTED_VALUE"
        prompt_for_option "Páginas a rotar (vacío = todas)" ""
        page_range="$PROMPTED_VALUE"
    fi
    for t in "${targets[@]}"; do
        rotate_pages "$t" "$angle" "$page_range" "${OUTPUT_PATH:-}"
    done
}

run_reorder() {
    local range="z-1" targets=()
    while [[ $# -gt 0 ]]; do
        case "$1" in
            --range) range="$2"; shift 2 ;;
            *) targets+=("$1"); shift ;;
        esac
    done
    if [[ ${#targets[@]} -eq 0 ]]; then
        prompt_for_file "Ruta del PDF"
        targets=("$PROMPTED_FILE")
        prompt_for_option "Rango de reordenamiento (z-1 = invertir)" "z-1"
        range="$PROMPTED_VALUE"
    fi
    for t in "${targets[@]}"; do
        reorder_pages "$t" "$range" "${OUTPUT_PATH:-}"
    done
}

run_delete() {
    local pages="" targets=()
    while [[ $# -gt 0 ]]; do
        case "$1" in
            --pages|-p) pages="$2"; shift 2 ;;
            *) targets+=("$1"); shift ;;
        esac
    done
    if [[ ${#targets[@]} -eq 0 || -z "$pages" ]]; then
        prompt_for_file "Ruta del PDF"
        targets=("$PROMPTED_FILE")
        prompt_for_option "Páginas a ELIMINAR (ej: 3,5,10-12)" ""
        pages="$PROMPTED_VALUE"
    fi
    for t in "${targets[@]}"; do
        delete_pages "$t" "$pages" "${OUTPUT_PATH:-}"
    done
}

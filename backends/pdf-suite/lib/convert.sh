#!/usr/bin/env bash
# =============================================================================
# lib/convert.sh — Conversiones PDF ↔ imagen ↔ texto ↔ HTML
# Motores: pdftoppm · pdftocairo · pdftotext · pdftohtml · img2pdf · convert
# =============================================================================

[[ -n "${_CONVERT_SOURCED:-}" ]] && return 0
readonly _CONVERT_SOURCED=1

# pdf_to_images()
# Convierte cada página de un PDF a una imagen (PNG, JPEG, SVG).
# Los archivos se guardan en un subdirectorio junto al PDF original.
#
# Arguments:
#   $1 - archivo PDF de entrada
#   $2 - formato de salida: png | jpg | svg (default: png)
#   $3 - DPI de renderizado (default: 150)
pdf_to_images() {
    local input="$1"
    local format="${2:-${RENDER_FORMAT}}"
    local dpi="${3:-${RENDER_DPI}}"

    validate_pdf_file "$input" || return 1

    local dir base out_dir
    dir="$(dirname "$input")"
    base="$(basename "$input" .pdf)"
    out_dir="${dir}/${base}_images"

    log_step "PDF → imágenes (${format}, ${dpi} DPI): $(basename "$input")"

    if [[ "${DRY_RUN:-false}" == "true" ]]; then
        log_skip "[dry-run] Crearía directorio: $(basename "$out_dir")"
        return 0
    fi

    mkdir -p "$out_dir" || {
        log_error "No se pudo crear el directorio: $out_dir"
        return 1
    }

    case "$format" in
        png)
            check_optional_dep "pdftoppm" "conversión a PNG" || return 1
            if pdftoppm -r "$dpi" -png "$input" "${out_dir}/${base}" 2>/dev/null; then
                local count
                count="$(find "$out_dir" -name "*.png" | wc -l)"
                log_success "${count} imágenes PNG en: $(basename "$out_dir")/"
            else
                log_failure "Error al convertir a PNG."
                return 1
            fi
            ;;
        jpg|jpeg)
            check_optional_dep "pdftoppm" "conversión a JPEG" || return 1
            if pdftoppm -r "$dpi" -jpeg "$input" "${out_dir}/${base}" 2>/dev/null; then
                local count
                count="$(find "$out_dir" -name "*.jpg" -o -name "*.jpeg" | wc -l)"
                log_success "${count} imágenes JPEG en: $(basename "$out_dir")/"
            else
                log_failure "Error al convertir a JPEG."
                return 1
            fi
            ;;
        svg)
            check_optional_dep "pdftocairo" "conversión a SVG" || return 1
            local total
            total="$(get_page_count "$input")"
            local ok=0
            for (( p=1; p<=total; p++ )); do
                local out_file
                out_file="$(printf "%s/%s_%03d.svg" "$out_dir" "$base" "$p")"
                pdftocairo -f "$p" -l "$p" -svg "$input" "$out_file" 2>/dev/null && (( ok++ ))
            done
            log_success "${ok}/${total} páginas SVG en: $(basename "$out_dir")/"
            ;;
        *)
            log_error "Formato no soportado: '${format}'. Usa: png, jpg, svg"
            return 1
            ;;
    esac
}

# pdf_to_text()
# Extrae el texto de un PDF preservando el layout si se pide.
#
# Arguments:
#   $1 - archivo PDF de entrada
#   $2 - (opcional) path de salida .txt
#   $3 - modo: plain | layout | table (default: plain)
pdf_to_text() {
    local input="$1"
    local output="${2:-}"
    local mode="${3:-plain}"

    check_optional_dep "pdftotext" "extracción de texto" || return 1
    validate_pdf_file "$input" || return 1

    local dir base
    dir="$(dirname "$input")"
    base="$(basename "$input" .pdf)"
    output="${output:-${dir}/${base}.txt}"

    validate_output_path "$output" || return 1

    log_step "PDF → texto (modo: ${mode}): $(basename "$input")"

    if [[ "${DRY_RUN:-false}" == "true" ]]; then
        log_skip "[dry-run] Crearía: $(basename "$output")"
        return 0
    fi

    local extra_args=()
    [[ "$mode" == "layout" ]] && extra_args+=(-layout)
    [[ "$mode" == "table"  ]] && extra_args+=(-table)

    if pdftotext "${extra_args[@]}" -enc UTF-8 "$input" "$output" 2>/dev/null; then
        local words
        words="$(wc -w < "$output")"
        log_success "$(basename "$output")  (${words} palabras)"
    else
        log_failure "Error al extraer texto de $(basename "$input")."
        rm -f "$output"
        return 1
    fi
}

# pdf_to_html()
# Convierte un PDF a HTML preservando estructura visual.
#
# Arguments:
#   $1 - archivo PDF de entrada
#   $2 - (opcional) path de salida (sin extensión, pdftohtml añade .html)
pdf_to_html() {
    local input="$1"
    local output_base="${2:-}"

    check_optional_dep "pdftohtml" "conversión a HTML" || return 1
    validate_pdf_file "$input" || return 1

    local dir base
    dir="$(dirname "$input")"
    base="$(basename "$input" .pdf)"
    output_base="${output_base:-${dir}/${base}}"

    log_step "PDF → HTML: $(basename "$input")"

    if [[ "${DRY_RUN:-false}" == "true" ]]; then
        log_skip "[dry-run] Crearía: ${base}.html"
        return 0
    fi

    if pdftohtml -s -enc UTF-8 "$input" "$output_base" 2>/dev/null; then
        log_success "${base}.html creado"
    else
        log_failure "Error al convertir a HTML."
        return 1
    fi
}

# images_to_pdf()
# Convierte una o más imágenes a un PDF sin recomprimir (usando img2pdf).
# Si img2pdf no está disponible, usa convert de ImageMagick como fallback.
#
# Arguments:
#   $1   - path de salida .pdf
#   $2..N - archivos de imagen en orden
images_to_pdf() {
    local output="$1"
    shift
    local images=("$@")

    if [[ ${#images[@]} -eq 0 ]]; then
        log_error "Se necesita al menos una imagen."
        return 1
    fi

    for img in "${images[@]}"; do
        if [[ ! -f "$img" ]]; then
            log_error "Imagen no encontrada: '${img}'"
            return 1
        fi
    done

    validate_output_path "$output" || return 1

    log_step "Imágenes → PDF: ${#images[@]} archivo(s) → $(basename "$output")"

    if [[ "${DRY_RUN:-false}" == "true" ]]; then
        log_skip "[dry-run] Crearía: $(basename "$output")"
        return 0
    fi

    if command -v img2pdf &>/dev/null; then
        if img2pdf "${images[@]}" -o "$output" 2>/dev/null; then
            local sz
            sz="$(format_file_size "$(get_file_size_bytes "$output")")"
            log_success "$(basename "$output") creado con img2pdf (sin pérdida, ${sz})"
            return 0
        fi
    fi

    # Fallback: ImageMagick convert
    if check_optional_dep "convert" "conversión de imágenes a PDF (fallback)" 2>/dev/null; then
        log_warn "img2pdf no disponible; usando ImageMagick (puede recomprimir)."
        if convert "${images[@]}" "$output" 2>/dev/null; then
            local sz
            sz="$(format_file_size "$(get_file_size_bytes "$output")")"
            log_success "$(basename "$output") creado con ImageMagick (${sz})"
        else
            log_failure "Error al convertir imágenes a PDF."
            return 1
        fi
    else
        log_error "Instala img2pdf o imagemagick para esta operación."
        return 1
    fi
}

# extract_images_from_pdf()
# Extrae todas las imágenes embebidas en un PDF.
#
# Arguments:
#   $1 - archivo PDF de entrada
#   $2 - formato: png | jpg | all (default: all)
extract_images_from_pdf() {
    local input="$1"
    local format="${2:-all}"

    check_optional_dep "pdfimages" "extracción de imágenes" || return 1
    validate_pdf_file "$input" || return 1

    local dir base out_dir
    dir="$(dirname "$input")"
    base="$(basename "$input" .pdf)"
    out_dir="${dir}/${base}_extracted_images"

    log_step "Extrayendo imágenes de $(basename "$input")"

    if [[ "${DRY_RUN:-false}" == "true" ]]; then
        log_skip "[dry-run] Crearía directorio: $(basename "$out_dir")"
        return 0
    fi

    mkdir -p "$out_dir"

    local flag
    case "$format" in
        png) flag="-png" ;;
        jpg) flag="-j"   ;;
        *)   flag="-all" ;;
    esac

    if pdfimages $flag "$input" "${out_dir}/${base}" 2>/dev/null; then
        local count
        count="$(find "$out_dir" -type f | wc -l)"
        log_success "${count} imágenes extraídas en: $(basename "$out_dir")/"
    else
        log_failure "Error al extraer imágenes."
        return 1
    fi
}

# run_convert()
# Punto de entrada del módulo de conversión.
run_convert() {
    local to_format="" dpi="$RENDER_DPI" targets=()
    while [[ $# -gt 0 ]]; do
        case "$1" in
            --to)      to_format="$2"; shift 2 ;;
            --dpi)     dpi="$2";       shift 2 ;;
            --extract-images) 
                shift
                for t in "$@"; do extract_images_from_pdf "$t"; done
                return ;;
            *) targets+=("$1"); shift ;;
        esac
    done

    if [[ ${#targets[@]} -eq 0 || -z "$to_format" ]]; then
        prompt_for_file "Ruta del PDF o imagen(es)"
        targets=("$PROMPTED_FILE")
        prompt_for_option "Convertir a (png|jpg|svg|txt|html)" "png"
        to_format="$PROMPTED_VALUE"
        [[ "$to_format" =~ ^(png|jpg|svg)$ ]] && {
            prompt_for_option "DPI" "$dpi"
            dpi="$PROMPTED_VALUE"
        }
    fi

    for t in "${targets[@]}"; do
        case "$to_format" in
            png|jpg|jpeg|svg) pdf_to_images "$t" "$to_format" "$dpi" ;;
            txt|text)         pdf_to_text   "$t" "${OUTPUT_PATH:-}" ;;
            html)             pdf_to_html   "$t" ;;
            pdf)              images_to_pdf "${OUTPUT_PATH:-${t%.*}.pdf}" "${targets[@]}"; break ;;
            *)
                log_error "Formato desconocido: '${to_format}'. Opciones: png jpg svg txt html"
                return 1 ;;
        esac
    done
}

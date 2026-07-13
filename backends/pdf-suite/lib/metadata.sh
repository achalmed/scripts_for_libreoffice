#!/usr/bin/env bash
# =============================================================================
# lib/metadata.sh — Lectura y escritura de metadatos PDF (XMP + DocInfo)
# Motores: exiftool (lectura completa + escritura) · pdfinfo (lectura rápida)
#          pdftk (dump_data / update_info para marcadores y formularios)
# =============================================================================
# CORRECCIÓN BUG #3 (arch_pdf_metadata_commands.sh):
#   qpdf --show-object no existe; se usa qpdf --check para validar estructura.
# CORRECCIÓN BUG #4 (arch_pdf_metadata_commands.sh):
#   mutool show trailer/Info apunta a DocInfo, NO a XMP.
#   Se documenta la distinción y se usan las rutas correctas.
# CORRECCIÓN BUG #5 (while+read sin IFS= -r):
#   Los bucles de find usan IFS= read -r para manejar espacios en nombres.
# =============================================================================

[[ -n "${_METADATA_SOURCED:-}" ]] && return 0
readonly _METADATA_SOURCED=1

# show_pdf_info()
# Muestra información completa de un PDF: páginas, tamaño, metadatos DocInfo y XMP.
# Combina pdfinfo (rápido) con exiftool (completo) si está disponible.
#
# Arguments:
#   $1 - archivo PDF
show_pdf_info() {
    local input="$1"

    validate_pdf_file "$input" || return 1

    local file_size total_pages
    file_size="$(format_file_size "$(get_file_size_bytes "$input")")"
    total_pages="$(get_page_count "$input")"

    log_section "Información: $(basename "$input")"
    echo -e "  ${C_DIM}Ruta completa:${C_RESET} $input"
    echo -e "  ${C_DIM}Tamaño:${C_RESET}        $file_size"
    echo -e "  ${C_DIM}Páginas:${C_RESET}       $total_pages"
    echo ""

    if check_optional_dep "pdfinfo" "información básica" 2>/dev/null; then
        echo -e "${C_BOLD}  — DocInfo (pdfinfo) —${C_RESET}"
        pdfinfo "$input" 2>/dev/null | grep -v "^File size\|^Pages:" | \
            sed 's/^/  /'
        echo ""
    fi

    if command -v exiftool &>/dev/null; then
        echo -e "${C_BOLD}  — Metadatos XMP completos (exiftool) —${C_RESET}"
        exiftool -XMP:all "$input" 2>/dev/null | sed 's/^/  /' | \
            grep -v "^\s*$"
        echo ""
    else
        log_warn "exiftool no está instalado; instálalo para ver metadatos XMP completos."
        log_warn "Instalar: sudo apt install libimage-exiftool-perl"
    fi
}

# read_metadata()
# Muestra todos los metadatos de uno o más PDFs en formato tabular.
# Usa exiftool -json para output estructurado si se pide.
#
# Arguments:
#   $1      - archivo PDF (o glob)
#   --json  - (flag) output en formato JSON
read_metadata() {
    local input="$1"
    local as_json=false
    [[ "${2:-}" == "--json" ]] && as_json=true

    check_optional_dep "exiftool" "lectura de metadatos" || return 1
    validate_pdf_file "$input" || return 1

    log_step "Metadatos de: $(basename "$input")"

    if [[ "$as_json" == "true" ]]; then
        exiftool -json "$input" 2>/dev/null
    else
        exiftool -Title -Author -Creator -Subject -Keywords \
                 -Description -CreateDate -ModifyDate \
                 -PDF:all "$input" 2>/dev/null | \
            grep -v "^\s*$" | sed 's/^/  /'
    fi
    echo ""
}

# write_metadata()
# Escribe metadatos en uno o varios PDFs usando exiftool.
# Modifica el archivo en-place (exiftool crea .original de backup por defecto).
#
# Arguments:
#   --title     "Título"
#   --author    "Autor"
#   --subject   "Asunto / Descripción corta"
#   --keywords  "kw1,kw2,kw3"
#   --creator   "Software creador"
#   --no-backup  no crear archivo .original (--overwrite_original)
#   $@          archivos PDF a modificar
write_metadata() {
    check_optional_dep "exiftool" "escritura de metadatos" || return 1

    local title="" author="" subject="" keywords="" creator=""
    local no_backup=false
    local targets=()

    while [[ $# -gt 0 ]]; do
        case "$1" in
            --title)    title="$2";    shift 2 ;;
            --author)   author="$2";   shift 2 ;;
            --subject)  subject="$2";  shift 2 ;;
            --keywords) keywords="$2"; shift 2 ;;
            --creator)  creator="$2";  shift 2 ;;
            --no-backup) no_backup=true; shift ;;
            *) targets+=("$1"); shift ;;
        esac
    done

    if [[ ${#targets[@]} -eq 0 ]]; then
        log_error "Debes indicar al menos un archivo PDF."
        return 1
    fi

    # Construir argumentos de exiftool
    local exif_args=()
    [[ -n "$title"    ]] && exif_args+=("-Title=${title}")
    [[ -n "$author"   ]] && exif_args+=("-Author=${author}")
    [[ -n "$subject"  ]] && exif_args+=("-Subject=${subject}")
    [[ -n "$keywords" ]] && exif_args+=("-Keywords=${keywords}")
    [[ -n "$creator"  ]] && exif_args+=("-Creator=${creator}")
    [[ "$no_backup" == "true" ]] && exif_args+=("-overwrite_original")

    if [[ ${#exif_args[@]} -le 1 ]]; then
        log_error "Debes especificar al menos un campo a escribir (--title, --author, etc.)"
        return 1
    fi

    for pdf in "${targets[@]}"; do
        validate_pdf_file "$pdf" || continue

        if [[ "${DRY_RUN:-false}" == "true" ]]; then
            log_skip "[dry-run] Actualizaría metadatos de: $(basename "$pdf")"
            continue
        fi

        log_step "Escribiendo metadatos en: $(basename "$pdf")"

        if exiftool "${exif_args[@]}" "$pdf" 2>/dev/null; then
            log_success "$(basename "$pdf") actualizado"
            [[ "$no_backup" != "true" ]] && \
                log_debug "Backup guardado en: ${pdf}_original"
        else
            log_failure "Error al escribir metadatos en $(basename "$pdf")."
        fi
    done
}

# set_title_from_filename()
# Establece el campo Title de cada PDF igual a su nombre de archivo (sin extensión).
# Útil para normalizar una colección de PDFs de forma masiva.
#
# Arguments:
#   $@ - archivos PDF; si es directorio, procesa todos los PDFs dentro
set_title_from_filename() {
    check_optional_dep "exiftool" "metadatos desde nombre de archivo" || return 1

    local targets=()
    for arg in "$@"; do
        if [[ -d "$arg" ]]; then
            while IFS= read -r f; do
                targets+=("$f")
            done < <(find "$arg" -type f -name "*.pdf" ${RECURSIVE:+-maxdepth 99} 2>/dev/null)
        else
            targets+=("$arg")
        fi
    done

    log_step "Estableciendo Title desde nombre de archivo (${#targets[@]} PDFs)"

    for pdf in "${targets[@]}"; do
        validate_pdf_file "$pdf" 2>/dev/null || continue
        local title
        title="$(basename "$pdf" .pdf)"

        if [[ "${DRY_RUN:-false}" == "true" ]]; then
            log_skip "[dry-run] Title='${title}' → $(basename "$pdf")"
            continue
        fi

        if exiftool -overwrite_original -Title="${title}" "$pdf" 2>/dev/null; then
            log_success "$(basename "$pdf") → Title='${title}'"
        else
            log_failure "Error en $(basename "$pdf")"
        fi
    done
}

# check_pdfs_without_metadata()
# Identifica PDFs en un directorio que no tienen el campo Title definido.
# Útil para detectar qué archivos faltan procesar con Zotero u otro flujo.
#
# Arguments:
#   $1 - directorio a escanear
check_pdfs_without_metadata() {
    local target_dir="${1:-$HOME/Zotero/storage}"

    check_optional_dep "exiftool" "escaneo de metadatos" || return 1

    if [[ ! -d "$target_dir" ]]; then
        log_error "Directorio no encontrado: '${target_dir}'"
        return 1
    fi

    log_section "PDFs sin metadatos en: $target_dir"
    local count_missing=0 count_ok=0

    while IFS= read -r pdf; do
        local title
        title="$(exiftool -s3 -Title "$pdf" 2>/dev/null)"
        if [[ -z "$title" ]]; then
            echo -e "  ${C_YELLOW}${SYM_WARN}${C_RESET} $pdf"
            (( count_missing++ ))
        else
            (( count_ok++ ))
        fi
    done < <(find "$target_dir" -name "*.pdf" -type f)

    echo ""
    log_summary_row "Con metadatos"   "$count_ok"      "${C_GREEN}"
    log_summary_row "Sin Title"       "$count_missing" "${C_YELLOW}"
    echo ""
}

# run_metadata()
# Punto de entrada del módulo de metadatos.
run_metadata() {
    local mode="read"
    local targets=()
    local write_args=()

    while [[ $# -gt 0 ]]; do
        case "$1" in
            --set-title)    mode="write"; write_args+=(--title "$2");    shift 2 ;;
            --set-author)   mode="write"; write_args+=(--author "$2");   shift 2 ;;
            --set-subject)  mode="write"; write_args+=(--subject "$2");  shift 2 ;;
            --set-keywords) mode="write"; write_args+=(--keywords "$2"); shift 2 ;;
            --set-creator)  mode="write"; write_args+=(--creator "$2");  shift 2 ;;
            --from-filename) mode="from-filename"; shift ;;
            --check-missing) mode="check-missing"; shift ;;
            --no-backup)    write_args+=(--no-backup); shift ;;
            --json)         mode="json"; shift ;;
            *) targets+=("$1"); shift ;;
        esac
    done

    case "$mode" in
        read)
            [[ ${#targets[@]} -eq 0 ]] && {
                prompt_for_file "Ruta del PDF"
                targets=("$PROMPTED_FILE")
            }
            for t in "${targets[@]}"; do show_pdf_info "$t"; done
            ;;
        json)
            for t in "${targets[@]}"; do read_metadata "$t" --json; done
            ;;
        write)
            [[ ${#targets[@]} -eq 0 ]] && {
                prompt_for_file "Ruta del PDF"
                targets=("$PROMPTED_FILE")
            }
            write_metadata "${write_args[@]}" "${targets[@]}"
            ;;
        from-filename)
            [[ ${#targets[@]} -eq 0 ]] && {
                prompt_for_file "Ruta del PDF o directorio"
                targets=("$PROMPTED_FILE")
            }
            set_title_from_filename "${targets[@]}"
            ;;
        check-missing)
            local dir="${targets[0]:-$HOME/Zotero/storage}"
            check_pdfs_without_metadata "$dir"
            ;;
    esac
}

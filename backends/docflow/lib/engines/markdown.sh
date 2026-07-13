#!/usr/bin/env bash
# lib/engines/markdown.sh — Markdown Engine.
# Convierte cualquier formato registrado a Markdown de alta calidad:
#   · pandoc para lo que pandoc lee bien (docx, odt, rtf, epub, html, tex)
#   · parsers OOXML propios para lo que pandoc NO lee (pptx, xlsx) —
#     preservan notas del presentador, tablas y todas las hojas
#   · LibreOffice como modernizador de formatos legados (doc→docx, ppt→pptx)
#   · pdftotext (+OCR opcional) para PDF
# Post-proceso común: imágenes ordenadas en <base>_files/images/, limpieza,
# frontmatter YAML con metadatos y sabor (flavor) de Markdown.

_MD_HELPERS="${DOCFLOW_ROOT}/lib/helpers"

# Contrato del dispatcher -----------------------------------------------------
engine_md_output_ext() {
    local strategy="${REG_MD_STRATEGY[$1]:-none}"
    [[ "$strategy" == "none" ]] && return 0
    printf 'md'
}

engine_md_convert() {
    local input="$1" output="$2"
    local ext strategy
    ext="$(util_ext "$input")"
    strategy="${REG_MD_STRATEGY[$ext]}"

    local media_dir="${output%.md}_files"
    rm -rf -- "$media_dir"

    case "$strategy" in
        pandoc)       _md_from_pandoc "$input" "$output" "$media_dir" ;;
        via_docx)     _md_via_intermediate "$input" "$output" "$media_dir" docx _md_from_pandoc ;;
        pptx_parser)  _md_from_pptx "$input" "$output" "$media_dir" ;;
        via_pptx)     _md_via_intermediate "$input" "$output" "$media_dir" pptx _md_from_pptx ;;
        xlsx_parser)  _md_from_xlsx "$input" "$output" ;;
        via_xlsx)     _md_via_intermediate "$input" "$output" "" xlsx _md_from_xlsx ;;
        csv_parser)   _md_from_csv "$input" "$output" ;;
        pdf_text)     _md_from_pdf "$input" "$output" ;;
        *)            log_error "Estrategia MD desconocida: $strategy"; return 1 ;;
    esac || return 1

    _md_postprocess "$input" "$output" "$media_dir"
}

# Estrategias ----------------------------------------------------------------

# _md_via_intermediate <in> <out> <media_dir> <ext_intermedia> <fn_final>
# Moderniza con LibreOffice y delega en la estrategia final.
_md_via_intermediate() {
    local input="$1" output="$2" media_dir="$3" mid_ext="$4" final_fn="$5"
    deps_require soffice
    local intermediate rc
    intermediate="$(soffice_intermediate "$input" "$mid_ext")" || {
        log_debug "LibreOffice no pudo modernizar: $input"
        return 1
    }
    "$final_fn" "$intermediate" "$output" "$media_dir"
    rc=$?
    rm -rf -- "$(dirname -- "$intermediate")"
    return $rc
}

_md_from_pandoc() {
    local input="$1" output="$2" media_dir="$3"
    deps_require pandoc
    local -a cmd=(pandoc "$input" -o "$output"
                  -t "$(_md_pandoc_target)" --wrap=none --markdown-headings=atx
                  --extract-media="$media_dir")
    # shellcheck disable=SC2206
    [[ -n "${DOCFLOW_PANDOC_ARGS:-}" ]] && cmd+=(${DOCFLOW_PANDOC_ARGS})
    "${cmd[@]}" 2> >(head -3 | while read -r l; do log_debug "pandoc: $l"; done)
}

# Variante de Markdown según el flavor pedido (los destinos tipo Obsidian,
# MkDocs, Hugo… consumen GFM sin extensiones raras de pandoc).
_md_pandoc_target() {
    case "${DOCFLOW_MD_FLAVOR:-gfm}" in
        quarto) printf 'markdown' ;;
        *)      printf 'gfm' ;;
    esac
}

_md_from_pptx() {
    local input="$1" output="$2" media_dir="$3"
    local -a extra=()
    [[ "${DOCFLOW_MD_INCLUDE_NOTES:-true}" == "true" ]] || extra+=(--no-notes)
    python3 "${_MD_HELPERS}/pptx_to_md.py" "$input" "$output" \
        --media-dir "$media_dir" \
        "--slide-separator=${DOCFLOW_MD_SLIDE_SEPARATOR:---}" \
        "${extra[@]}"
}

_md_from_xlsx() {
    local input="$1" output="$2"
    python3 "${_MD_HELPERS}/xlsx_to_md.py" "$input" "$output"
}

_md_from_csv() {
    local input="$1" output="$2"
    python3 "${_MD_HELPERS}/xlsx_to_md.py" "$input" "$output" --csv
}

_md_from_pdf() {
    local input="$1" output="$2"
    deps_require pdftotext
    local source_pdf="$input" tmp=""

    # OCR previo opcional para PDFs escaneados
    if [[ "${DOCFLOW_OCR:-false}" == "true" ]] && deps_has ocrmypdf; then
        tmp="$(util_mktemp_dir ocr)/ocr.pdf"
        if ocrmypdf --skip-text -l "${DOCFLOW_OCR_LANG:-spa+eng}" \
                "$input" "$tmp" >/dev/null 2>&1; then
            source_pdf="$tmp"
        fi
    fi

    local txt rc=1
    txt="$(mktemp "${TMPDIR:-/tmp}/docflow_pdf_XXXXXX.txt")"
    if pdftotext -layout -enc UTF-8 "$source_pdf" "$txt" 2>/dev/null; then
        {
            printf '# %s\n\n' "$(util_basename_noext "$input")"
            # Colapsar 3+ líneas en blanco y eliminar saltos de página
            tr -d '\f' < "$txt" | cat -s
        } > "$output"
        rc=0
    fi
    rm -f -- "$txt"
    [[ -n "$tmp" ]] && rm -rf -- "$(dirname -- "$tmp")"
    return $rc
}

# Post-proceso común ----------------------------------------------------------

# _md_postprocess <input_original> <output_md> <media_dir>
_md_postprocess() {
    local input="$1" output="$2" media_dir="$3"

    # 1. Imágenes: renombrar a figure-NNN, mover a <base>_files/images/ y
    #    reescribir los enlaces como rutas relativas.
    if [[ -n "$media_dir" && -d "$media_dir" ]]; then
        python3 "${_MD_HELPERS}/mdtools.py" fix-media "$output" --media-dir "$media_dir" || return 1
        media_process_dir "${media_dir}/images" "$output"
    fi

    # 2. Limpieza: HTML residual, atributos basura, líneas en blanco duplicadas.
    if [[ "${DOCFLOW_MD_CLEAN:-true}" == "true" ]]; then
        python3 "${_MD_HELPERS}/mdtools.py" clean "$output" || return 1
    fi

    # 3. Frontmatter YAML con los metadatos del documento original.
    if [[ "${DOCFLOW_MD_METADATA:-true}" == "true" ]]; then
        metadata_prepend_frontmatter "$input" "$output"
    fi
    return 0
}

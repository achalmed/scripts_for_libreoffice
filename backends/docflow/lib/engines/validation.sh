#!/usr/bin/env bash
# lib/engines/validation.sh — Validation Engine.
# Verificación automática post-conversión: la salida existe, no está vacía
# y es estructuralmente válida. Si la verificación falla, el archivo original
# jamás se marca como convertido (⇒ nunca se elimina/respalda).

# validation_verify <task> <output> → 0 si la salida es válida
validation_verify() {
    local task="$1" output="$2"

    [[ -s "$output" ]] || { log_debug "verificación: salida vacía o ausente: $output"; return 1; }

    case "$task" in
        md)     _verify_markdown "$output" ;;
        pdf)    _verify_pdf "$output" ;;
        odf|office) _verify_zip_container "$output" ;;
        *)      return 0 ;;
    esac
}

# Markdown: verifica que las imágenes referenciadas existen y que el
# contenido no quedó vacío tras el frontmatter.
_verify_markdown() {
    python3 "${DOCFLOW_ROOT}/lib/helpers/mdtools.py" verify "$1"
}

# PDF: qpdf --check si está disponible; si no, firma %PDF- en cabecera.
_verify_pdf() {
    local pdf="$1"
    if deps_has qpdf; then
        qpdf --check "$pdf" >/dev/null 2>&1
    else
        head -c 5 -- "$pdf" | grep -q '^%PDF-'
    fi
}

# ODF/OOXML: deben ser contenedores ZIP válidos (los planos .fodt/.fods/.fodp
# y los legados .doc/.xls/.ppt solo se comprueban como no-vacíos).
_verify_zip_container() {
    local file="$1" ext
    ext="$(util_ext "$file")"
    case "$ext" in
        fodt|fods|fodp|doc|xls|ppt|dot|pot|rtf|txt|csv|tsv) return 0 ;;
        *) util_is_zip "$file" ;;
    esac
}

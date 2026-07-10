#!/usr/bin/env bash
# lib/engines/pdf_ops.sh — Operaciones sobre PDFs existentes.
# Unir, dividir, rotar, extraer páginas, comprimir, PDF/A, cifrar/descifrar,
# marca de agua, numeración y OCR. Herramientas: qpdf (estructura),
# ghostscript (re-render/compresión/PDF/A), ocrmypdf (OCR).
# Cada operación escribe de forma segura: primero a temporal, luego mv.

# _pdf_ops_safe_out <destino> — imprime ruta temporal junto al destino
_pdf_ops_safe_out() {
    mktemp "$(dirname -- "$1")/.docflow_pdf_XXXXXX.pdf"
}

# pdf_ops_merge <salida> <entrada...>
pdf_ops_merge() {
    deps_require qpdf
    local out="$1"; shift
    qpdf --empty --pages "$@" -- "$out" && [[ -s "$out" ]]
}

# pdf_ops_split <entrada> [páginas_por_trozo]
# Genera <base>-NN.pdf junto al original.
pdf_ops_split() {
    deps_require qpdf
    local input="$1" every="${2:-1}"
    local base_out
    base_out="$(dirname -- "$input")/$(util_basename_noext "$input")"
    qpdf --split-pages="$every" "$input" "${base_out}-%d.pdf"
}

# pdf_ops_rotate <entrada> <salida> <grados> [rango_páginas]
pdf_ops_rotate() {
    deps_require qpdf
    local input="$1" out="$2" degrees="$3" pages="${4:-1-z}"
    local tmp; tmp="$(_pdf_ops_safe_out "$out")"
    qpdf "$input" --rotate="+${degrees}:${pages}" -- "$tmp" && mv -- "$tmp" "$out" || { rm -f -- "$tmp"; return 1; }
}

# pdf_ops_extract <entrada> <salida> <rango_páginas>  (ej: 1-5,9,12-z)
pdf_ops_extract() {
    deps_require qpdf
    local input="$1" out="$2" pages="$3"
    local tmp; tmp="$(_pdf_ops_safe_out "$out")"
    qpdf --empty --pages "$input" "$pages" -- "$tmp" && mv -- "$tmp" "$out" || { rm -f -- "$tmp"; return 1; }
}

# pdf_ops_compress <entrada> <salida>  (ghostscript /ebook: buen equilibrio)
pdf_ops_compress() {
    deps_require gs
    local input="$1" out="$2"
    local tmp; tmp="$(_pdf_ops_safe_out "$out")"
    if gs -q -sDEVICE=pdfwrite -dCompatibilityLevel=1.5 -dPDFSETTINGS=/ebook \
          -dNOPAUSE -dBATCH -sOutputFile="$tmp" "$input" 2>/dev/null && [[ -s "$tmp" ]]; then
        # Solo reemplazar si de verdad redujo el tamaño
        if (( $(util_file_size "$tmp") < $(util_file_size "$input") )); then
            mv -- "$tmp" "$out"
        else
            [[ "$input" != "$out" ]] && cp -- "$input" "$out"
            rm -f -- "$tmp"
            log_info "El PDF ya estaba optimizado: $(basename -- "$input")"
        fi
        return 0
    fi
    rm -f -- "$tmp"; return 1
}

# pdf_ops_pdfa <entrada> <salida> — conversión a PDF/A-2b
pdf_ops_pdfa() {
    deps_require gs
    local input="$1" out="$2"
    local tmp; tmp="$(_pdf_ops_safe_out "$out")"
    gs -q -dPDFA=2 -dBATCH -dNOPAUSE -sColorConversionStrategy=UseDeviceIndependentColor \
       -sDEVICE=pdfwrite -dPDFACompatibilityPolicy=1 \
       -sOutputFile="$tmp" "$input" 2>/dev/null \
        && [[ -s "$tmp" ]] && mv -- "$tmp" "$out" || { rm -f -- "$tmp"; return 1; }
}

# pdf_ops_encrypt <entrada> <salida> <contraseña>
pdf_ops_encrypt() {
    deps_require qpdf
    local tmp; tmp="$(_pdf_ops_safe_out "$2")"
    qpdf --encrypt "$3" "$3" 256 -- "$1" "$tmp" && mv -- "$tmp" "$2" || { rm -f -- "$tmp"; return 1; }
}

# pdf_ops_decrypt <entrada> <salida> <contraseña>
pdf_ops_decrypt() {
    deps_require qpdf
    local tmp; tmp="$(_pdf_ops_safe_out "$2")"
    qpdf --password="$3" --decrypt -- "$1" "$tmp" && mv -- "$tmp" "$2" || { rm -f -- "$tmp"; return 1; }
}

# pdf_ops_watermark <entrada> <salida> <texto>
# Marca de agua diagonal generada al vuelo con ghostscript.
pdf_ops_watermark() {
    deps_require gs qpdf
    local input="$1" out="$2" text="$3"
    local stamp tmp
    stamp="$(mktemp "${TMPDIR:-/tmp}/docflow_wm_XXXXXX.pdf")"
    tmp="$(_pdf_ops_safe_out "$out")"

    gs -q -sDEVICE=pdfwrite -dBATCH -dNOPAUSE -sOutputFile="$stamp" \
       -dDEVICEWIDTHPOINTS=612 -dDEVICEHEIGHTPOINTS=792 \
       -c "<</EndPage {
             2 eq { pop false }
             { gsave 0.85 setgray /Helvetica-Bold 48 selectfont
               306 396 moveto 45 rotate ($text) dup stringwidth pop 2 div neg 0 rmoveto show
               grestore true } ifelse
           }>> setpagedevice" \
       -f "$input" 2>/dev/null

    if [[ -s "$stamp" ]]; then
        mv -- "$stamp" "$tmp" && mv -- "$tmp" "$out"
        return 0
    fi
    rm -f -- "$stamp" "$tmp"
    return 1
}

# pdf_ops_ocr <entrada> <salida> — capa de texto OCR sobre PDF escaneado
pdf_ops_ocr() {
    deps_require ocrmypdf
    ocrmypdf --skip-text -l "${DOCFLOW_OCR_LANG:-spa+eng}" "$1" "$2" >/dev/null 2>&1
}

# pdf_ops_info <entrada> — metadatos y propiedades del PDF
pdf_ops_info() {
    local input="$1"
    if deps_has pdftotext && command -v pdfinfo &>/dev/null; then
        pdfinfo -- "$input"
    elif deps_has exiftool; then
        exiftool -- "$input"
    elif deps_has qpdf; then
        qpdf --show-npages "$input" | xargs -I{} echo "Páginas: {}"
    else
        deps_require pdftotext
    fi
    if deps_has qpdf; then
        [[ "$(qpdf --is-encrypted "$input" 2>/dev/null; echo $?)" == "0" ]] \
            && log_warn "Este PDF está cifrado."
    fi
}

#!/usr/bin/env bash
# lib/commands/cmd_pdf.sh — Comando pdf: operaciones sobre PDFs existentes.
# Delegación 1:1 en el motor pdf_ops; aquí solo parsing de la sub-operación.

cmd_pdf() {
    local op="${1:-}"
    shift || true

    case "$op" in
        merge)     _cmd_pdf_merge "$@" ;;
        split)     _cmd_pdf_binary pdf_ops_split "dividir" "$@" ;;
        rotate)    _cmd_pdf_rotate "$@" ;;
        extract)   _cmd_pdf_extract "$@" ;;
        compress)  _cmd_pdf_inout pdf_ops_compress "comprimido" "$@" ;;
        pdfa)      _cmd_pdf_inout pdf_ops_pdfa "PDF/A" "$@" ;;
        encrypt)   _cmd_pdf_password pdf_ops_encrypt "cifrado" "$@" ;;
        decrypt)   _cmd_pdf_password pdf_ops_decrypt "descifrado" "$@" ;;
        watermark) _cmd_pdf_watermark "$@" ;;
        ocr)       _cmd_pdf_inout pdf_ops_ocr "OCR" "$@" ;;
        info)      [[ -f "${1:-}" ]] || { log_error "Uso: docflow pdf info <archivo.pdf>"; exit "$EXIT_USAGE"; }
                   pdf_ops_info "$1"; metadata_show "$1"; exit 0 ;;
        ""|-h|--help) cli_help_pdf; exit 0 ;;
        *) log_error "Operación PDF desconocida: '$op'"; cli_help_pdf; exit "$EXIT_USAGE" ;;
    esac
}

_cmd_pdf_merge() {
    [[ $# -lt 3 ]] && { log_error "Uso: docflow pdf merge <salida.pdf> <in1.pdf> <in2.pdf>..."; exit "$EXIT_USAGE"; }
    local out="$1"; shift
    if pdf_ops_merge "$out" "$@"; then
        log_ok "unido: ${out} ($# archivos)"
    else
        log_fail "no se pudo unir"; exit "$EXIT_CONVERT_FAIL"
    fi
}

# split y similares: <entrada> [args...]
_cmd_pdf_binary() {
    local fn="$1" verb="$2"; shift 2
    [[ -f "${1:-}" ]] || { log_error "Uso: docflow pdf ${verb} <archivo.pdf> [páginas_por_trozo]"; exit "$EXIT_USAGE"; }
    if "$fn" "$@"; then log_ok "${verb}: $1"; else log_fail "$verb"; exit "$EXIT_CONVERT_FAIL"; fi
}

# Operaciones <entrada> [-o salida]: por defecto sobreescriben con seguridad
_cmd_pdf_inout() {
    local fn="$1" verb="$2"; shift 2
    local input="" output=""
    while [[ $# -gt 0 ]]; do
        case "$1" in
            -o|--output) output="$2"; shift 2 ;;
            *) input="$1"; shift ;;
        esac
    done
    [[ -f "$input" ]] || { log_error "Uso: docflow pdf <op> <archivo.pdf> [-o salida.pdf]"; exit "$EXIT_USAGE"; }
    : "${output:=$input}"
    if "$fn" "$input" "$output"; then
        log_ok "${verb}: ${output} ($(util_human_size "$(util_file_size "$output")"))"
    else
        log_fail "$verb"; exit "$EXIT_CONVERT_FAIL"
    fi
}

_cmd_pdf_password() {
    local fn="$1" verb="$2"; shift 2
    local input="" output="" password=""
    while [[ $# -gt 0 ]]; do
        case "$1" in
            -o|--output) output="$2"; shift 2 ;;
            --password)  password="$2"; shift 2 ;;
            *) input="$1"; shift ;;
        esac
    done
    [[ -f "$input" && -n "$password" ]] || {
        log_error "Uso: docflow pdf ${verb%ado} <archivo.pdf> --password <clave> [-o salida.pdf]"
        exit "$EXIT_USAGE"
    }
    : "${output:=$input}"
    if "$fn" "$input" "$output" "$password"; then
        log_ok "${verb}: ${output}"
    else
        log_fail "${verb} falló (¿contraseña incorrecta?)"; exit "$EXIT_CONVERT_FAIL"
    fi
}

_cmd_pdf_rotate() {
    local input="" output="" degrees="90" pages="1-z"
    while [[ $# -gt 0 ]]; do
        case "$1" in
            --degrees) degrees="$2"; shift 2 ;;
            --pages)   pages="$2"; shift 2 ;;
            -o|--output) output="$2"; shift 2 ;;
            *) input="$1"; shift ;;
        esac
    done
    [[ -f "$input" ]] || { log_error "Uso: docflow pdf rotate <archivo.pdf> [--degrees 90] [--pages 1-3]"; exit "$EXIT_USAGE"; }
    : "${output:=$input}"
    if pdf_ops_rotate "$input" "$output" "$degrees" "$pages"; then
        log_ok "rotado ${degrees}°: ${output}"
    else
        log_fail "rotación"; exit "$EXIT_CONVERT_FAIL"
    fi
}

_cmd_pdf_extract() {
    local input="" output="" pages=""
    while [[ $# -gt 0 ]]; do
        case "$1" in
            --pages) pages="$2"; shift 2 ;;
            -o|--output) output="$2"; shift 2 ;;
            *) input="$1"; shift ;;
        esac
    done
    [[ -f "$input" && -n "$pages" && -n "$output" ]] || {
        log_error "Uso: docflow pdf extract <archivo.pdf> --pages 1-5 -o salida.pdf"
        exit "$EXIT_USAGE"
    }
    if pdf_ops_extract "$input" "$output" "$pages"; then
        log_ok "extraído (${pages}): ${output}"
    else
        log_fail "extracción"; exit "$EXIT_CONVERT_FAIL"
    fi
}

_cmd_pdf_watermark() {
    local input="" output="" text=""
    while [[ $# -gt 0 ]]; do
        case "$1" in
            --text) text="$2"; shift 2 ;;
            -o|--output) output="$2"; shift 2 ;;
            *) input="$1"; shift ;;
        esac
    done
    [[ -f "$input" && -n "$text" ]] || {
        log_error "Uso: docflow pdf watermark <archivo.pdf> --text \"BORRADOR\" [-o salida.pdf]"
        exit "$EXIT_USAGE"
    }
    : "${output:=$input}"
    if pdf_ops_watermark "$input" "$output" "$text"; then
        log_ok "marca de agua: ${output}"
    else
        log_fail "marca de agua"; exit "$EXIT_CONVERT_FAIL"
    fi
}

cli_help_pdf() {
    cat >&2 <<EOF
${C_BOLD}docflow pdf${C_RESET} — Operaciones sobre PDFs existentes.

USO: docflow pdf <operación> [argumentos]

OPERACIONES:
  merge <salida> <in...>              Unir varios PDFs
  split <archivo> [n]                 Dividir (n páginas por trozo, defecto 1)
  rotate <archivo> [--degrees 90] [--pages 1-3] [-o salida]
  extract <archivo> --pages 1-5,9 -o salida
  compress <archivo> [-o salida]      Comprimir/optimizar (ghostscript)
  pdfa <archivo> [-o salida]          Convertir a PDF/A-2b (archivado)
  encrypt <archivo> --password <p>    Proteger con contraseña (AES-256)
  decrypt <archivo> --password <p>    Quitar contraseña
  watermark <archivo> --text "..."    Marca de agua diagonal
  ocr <archivo> [-o salida]           Capa de texto OCR (ocrmypdf)
  info <archivo>                      Metadatos y propiedades

EJEMPLOS:
  docflow pdf merge tesis.pdf cap1.pdf cap2.pdf cap3.pdf
  docflow pdf compress escaneo.pdf -o escaneo_min.pdf
  docflow pdf watermark borrador.pdf --text "BORRADOR"
EOF
}

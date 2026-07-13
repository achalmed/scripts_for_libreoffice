#!/usr/bin/env bash
# =============================================================================
# lib/protect.sh — Protección de PDFs: cifrado, descifrado, marca de agua
# Motores: qpdf (cifrado AES-256) · cpdf (watermark texto) · pdftk (sello PDF)
# =============================================================================

[[ -n "${_PROTECT_SOURCED:-}" ]] && return 0
readonly _PROTECT_SOURCED=1

# encrypt_pdf()
# Cifra un PDF con contraseña usando AES-256 (estándar PDF 1.7+).
# La contraseña de usuario se necesita para abrir el archivo.
# La contraseña owner se necesita para modificar permisos.
#
# Arguments:
#   $1 - archivo de entrada
#   $2 - contraseña de usuario (vacía = sin contraseña para abrir)
#   $3 - contraseña owner
#   $4 - (opcional) path de salida
encrypt_pdf() {
    local input="$1"
    local user_pass="${2:-}"
    local owner_pass="${3:-}"
    local output="${4:-}"

    if [[ -z "$owner_pass" ]]; then
        log_error "La contraseña owner es obligatoria para cifrar."
        return 1
    fi

    validate_pdf_file "$input" || return 1

    local dir base
    dir="$(dirname "$input")"
    base="$(basename "$input" .pdf)"
    output="${output:-${dir}/${base}_protected.pdf}"

    validate_output_path "$output" || return 1

    log_step "Cifrando: $(basename "$input")"

    if [[ "${DRY_RUN:-false}" == "true" ]]; then
        log_skip "[dry-run] Cifraría: $(basename "$input")"
        return 0
    fi

    if qpdf --encrypt "$user_pass" "$owner_pass" 256 -- "$input" "$output" 2>/dev/null; then
        log_success "$(basename "$output") cifrado con AES-256"
    else
        log_failure "Error al cifrar $(basename "$input")."
        rm -f "$output"
        return 1
    fi
}

# decrypt_pdf()
# Descifra un PDF protegido con contraseña.
#
# Arguments:
#   $1 - archivo cifrado de entrada
#   $2 - contraseña (user o owner)
#   $3 - (opcional) path de salida
decrypt_pdf() {
    local input="$1"
    local password="${2:-}"
    local output="${3:-}"

    validate_pdf_file "$input" || return 1

    local dir base
    dir="$(dirname "$input")"
    base="$(basename "$input" .pdf)"
    output="${output:-${dir}/${base}_decrypted.pdf}"

    validate_output_path "$output" || return 1

    log_step "Descifrando: $(basename "$input")"

    if [[ "${DRY_RUN:-false}" == "true" ]]; then
        log_skip "[dry-run] Descifraría: $(basename "$input")"
        return 0
    fi

    if qpdf --password="$password" --decrypt "$input" "$output" 2>/dev/null; then
        log_success "$(basename "$output") descifrado correctamente"
    else
        log_failure "Error al descifrar. Verifica la contraseña."
        rm -f "$output"
        return 1
    fi
}

# add_text_watermark()
# Añade una marca de agua de texto diagonal usando cpdf.
# Si cpdf no está disponible, genera instrucciones de instalación.
#
# Arguments:
#   $1 - archivo de entrada
#   $2 - texto de la marca de agua (ej: "BORRADOR")
#   $3 - opacidad 0.0-1.0 (default: 0.3)
#   $4 - ángulo de rotación del texto (default: 45)
#   $5 - color RGB normalizado (default: "0.8 0 0" = rojo)
#   $6 - (opcional) path de salida
add_text_watermark() {
    local input="$1"
    local text="${2:-BORRADOR}"
    local opacity="${3:-0.3}"
    local angle="${4:-45}"
    local color="${5:-0.8 0 0}"
    local output="${6:-}"

    check_optional_dep "cpdf" "marca de agua de texto" || {
        echo -e "${C_YELLOW}  Para instalar cpdf:${C_RESET}"
        echo -e "  wget https://github.com/coherentgraphics/cpdf-binaries/raw/master/Linux-Intel/cpdf"
        echo -e "  chmod +x cpdf && sudo mv cpdf /usr/local/bin/"
        return 1
    }

    validate_pdf_file "$input" || return 1

    local dir base
    dir="$(dirname "$input")"
    base="$(basename "$input" .pdf)"
    output="${output:-${dir}/${base}_watermark.pdf}"

    validate_output_path "$output" || return 1

    log_step "Añadiendo marca de agua '${text}' a $(basename "$input")"

    if [[ "${DRY_RUN:-false}" == "true" ]]; then
        log_skip "[dry-run] Crearía: $(basename "$output")"
        return 0
    fi

    # shellcheck disable=SC2086
    if cpdf -add-text "$text" \
            -font "Helvetica-Bold" \
            -font-size 72 \
            -opacity "$opacity" \
            -rotate-font "$angle" \
            -color $color \
            -underneath \
            "$input" -o "$output" 2>/dev/null; then
        log_success "$(basename "$output") con marca de agua '${text}'"
    else
        log_failure "Error al añadir marca de agua."
        rm -f "$output"
        return 1
    fi
}

# add_pdf_stamp()
# Añade un sello PDF como capa encima del contenido (ej: logo, firma).
# Motor: pdftk stamp (sello encima) o background (sello debajo).
#
# Arguments:
#   $1 - archivo de entrada
#   $2 - archivo PDF del sello/logo
#   $3 - posición: foreground | background (default: foreground)
#   $4 - (opcional) path de salida
add_pdf_stamp() {
    local input="$1"
    local stamp_pdf="$2"
    local position="${3:-foreground}"
    local output="${4:-}"

    check_optional_dep "pdftk" "sello PDF" || return 1
    validate_pdf_file "$input" || return 1
    validate_pdf_file "$stamp_pdf" || return 1

    local dir base
    dir="$(dirname "$input")"
    base="$(basename "$input" .pdf)"
    output="${output:-${dir}/${base}_stamped.pdf}"

    validate_output_path "$output" || return 1

    log_step "Añadiendo sello $(basename "$stamp_pdf") a $(basename "$input")"

    if [[ "${DRY_RUN:-false}" == "true" ]]; then
        log_skip "[dry-run] Crearía: $(basename "$output")"
        return 0
    fi

    local pdftk_op
    [[ "$position" == "background" ]] && pdftk_op="background" || pdftk_op="stamp"

    if pdftk "$input" $pdftk_op "$stamp_pdf" output "$output" 2>/dev/null; then
        log_success "$(basename "$output") con sello aplicado"
    else
        log_failure "Error al aplicar el sello."
        rm -f "$output"
        return 1
    fi
}

# add_page_numbers()
# Añade números de página al pie de cada hoja usando cpdf.
#
# Arguments:
#   $1 - archivo de entrada
#   $2 - (opcional) path de salida
add_page_numbers() {
    local input="$1"
    local output="${2:-}"

    check_optional_dep "cpdf" "numeración de páginas" || return 1
    validate_pdf_file "$input" || return 1

    local dir base total
    dir="$(dirname "$input")"
    base="$(basename "$input" .pdf)"
    total="$(get_page_count "$input")"
    output="${output:-${dir}/${base}_numbered.pdf}"

    validate_output_path "$output" || return 1

    log_step "Añadiendo números de página a $(basename "$input") (${total} págs)"

    if [[ "${DRY_RUN:-false}" == "true" ]]; then
        log_skip "[dry-run] Crearía: $(basename "$output")"
        return 0
    fi

    if cpdf -add-text "Página %Page de %EndPage" \
            -font "Helvetica" \
            -font-size 9 \
            -bottom 8mm \
            -midline \
            "$input" -o "$output" 2>/dev/null; then
        log_success "$(basename "$output") con números de página"
    else
        log_failure "Error al numerar páginas."
        rm -f "$output"
        return 1
    fi
}

# run_protect()
# Punto de entrada del módulo de protección.
run_protect() {
    local mode=""
    local user_pass="" owner_pass="" password=""
    local wm_text="BORRADOR" wm_opacity="0.3" wm_angle="45" wm_color="0.8 0 0"
    local stamp_pdf="" stamp_pos="foreground"
    local targets=()

    while [[ $# -gt 0 ]]; do
        case "$1" in
            --encrypt)    mode="encrypt"; shift ;;
            --decrypt)    mode="decrypt"; shift ;;
            --watermark)  mode="watermark"; shift ;;
            --stamp)      mode="stamp"; stamp_pdf="$2"; shift 2 ;;
            --page-numbers) mode="page-numbers"; shift ;;
            --user-pass)  user_pass="$2";  shift 2 ;;
            --owner-pass) owner_pass="$2"; shift 2 ;;
            --password)   password="$2";   shift 2 ;;
            --text)       wm_text="$2";    shift 2 ;;
            --opacity)    wm_opacity="$2"; shift 2 ;;
            --angle)      wm_angle="$2";   shift 2 ;;
            --color)      wm_color="$2";   shift 2 ;;
            --background) stamp_pos="background"; shift ;;
            *) targets+=("$1"); shift ;;
        esac
    done

    if [[ -z "$mode" || ${#targets[@]} -eq 0 ]]; then
        if [[ -z "$mode" ]]; then
            prompt_for_option "Operación (encrypt|decrypt|watermark|stamp|page-numbers)" "encrypt"
            mode="$PROMPTED_VALUE"
        fi
        if [[ ${#targets[@]} -eq 0 ]]; then
            prompt_for_file "Ruta del PDF"
            targets=("$PROMPTED_FILE")
        fi
    fi

    for t in "${targets[@]}"; do
        case "$mode" in
            encrypt)
                [[ -z "$owner_pass" ]] && {
                    read -rsp "  Contraseña owner: " owner_pass; echo
                }
                encrypt_pdf "$t" "$user_pass" "$owner_pass" "${OUTPUT_PATH:-}"
                ;;
            decrypt)
                [[ -z "$password" ]] && {
                    read -rsp "  Contraseña: " password; echo
                }
                decrypt_pdf "$t" "$password" "${OUTPUT_PATH:-}"
                ;;
            watermark)
                add_text_watermark "$t" "$wm_text" "$wm_opacity" "$wm_angle" "$wm_color" "${OUTPUT_PATH:-}"
                ;;
            stamp)
                add_pdf_stamp "$t" "$stamp_pdf" "$stamp_pos" "${OUTPUT_PATH:-}"
                ;;
            page-numbers)
                add_page_numbers "$t" "${OUTPUT_PATH:-}"
                ;;
        esac
    done
}

# run_watermark() — alias para run_protect --watermark
run_watermark() {
    run_protect --watermark "$@"
}

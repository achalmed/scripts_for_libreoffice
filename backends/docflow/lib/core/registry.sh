#!/usr/bin/env bash
# lib/core/registry.sh — Registro central de formatos soportados.
# Patrón registry: cada archivo de lib/formats/ declara sus capacidades aquí.
# Añadir un formato nuevo NO requiere tocar el núcleo ni los motores:
# basta crear lib/formats/<ext>.sh con una llamada a registry_add.

# Capacidades por extensión (asociativos indexados por extensión):
declare -gA REG_CATEGORY=()      # document|presentation|spreadsheet|markup|data|pdf
declare -gA REG_DESCRIPTION=()   # descripción legible
declare -gA REG_MD_STRATEGY=()   # estrategia del Markdown Engine
declare -gA REG_PDF_STRATEGY=()  # estrategia del PDF Engine
declare -gA REG_ODF_TARGET=()    # extensión destino en OpenDocument ('' = n/a)
declare -gA REG_MS_TARGET=()     # extensión destino en MS Office ('' = n/a)

# registry_add <ext> <categoría> <descripción> <md_strategy> <pdf_strategy> \
#              <odf_target> <ms_target>
#
# Estrategias MD:  pandoc | via_docx | pptx_parser | via_pptx | xlsx_parser |
#                  via_xlsx | csv_parser | pdf_text | none
# Estrategias PDF: soffice | pandoc | pandoc_soffice | copy | none
registry_add() {
    local ext="$1"
    REG_CATEGORY[$ext]="$2"
    REG_DESCRIPTION[$ext]="$3"
    REG_MD_STRATEGY[$ext]="$4"
    REG_PDF_STRATEGY[$ext]="$5"
    REG_ODF_TARGET[$ext]="$6"
    REG_MS_TARGET[$ext]="$7"
}

# registry_supported <ext> → 0 si la extensión está registrada
registry_supported() {
    [[ -n "${REG_CATEGORY[$1]:-}" ]]
}

# registry_exts_for <campo> → lista de extensiones con ese campo no vacío.
# campo ∈ md|pdf|odf|ms|all
registry_exts_for() {
    local field="$1" ext out=()
    for ext in "${!REG_CATEGORY[@]}"; do
        case "$field" in
            all) out+=("$ext") ;;
            md)  [[ "${REG_MD_STRATEGY[$ext]}"  != "none" ]] && out+=("$ext") ;;
            pdf) [[ "${REG_PDF_STRATEGY[$ext]}" != "none" ]] && out+=("$ext") ;;
            odf) [[ -n "${REG_ODF_TARGET[$ext]}" ]] && out+=("$ext") ;;
            ms)  [[ -n "${REG_MS_TARGET[$ext]}"  ]] && out+=("$ext") ;;
        esac
    done
    printf '%s\n' "${out[@]}" | sort
}

# registry_print — tabla legible de formatos (comando `docflow formats`)
registry_print() {
    local ext
    printf '%-6s %-13s %-38s %-4s %-4s %-4s %-4s\n' \
        "EXT" "CATEGORÍA" "DESCRIPCIÓN" "MD" "PDF" "ODF" "MS"
    for ext in $(printf '%s\n' "${!REG_CATEGORY[@]}" | sort); do
        printf '%-6s %-13s %-38s %-4s %-4s %-4s %-4s\n' \
            "$ext" "${REG_CATEGORY[$ext]}" "${REG_DESCRIPTION[$ext]}" \
            "$([[ "${REG_MD_STRATEGY[$ext]}"  != "none" ]] && echo "✓" || echo "-")" \
            "$([[ "${REG_PDF_STRATEGY[$ext]}" != "none" ]] && echo "✓" || echo "-")" \
            "$([[ -n "${REG_ODF_TARGET[$ext]}" ]] && echo "✓" || echo "-")" \
            "$([[ -n "${REG_MS_TARGET[$ext]}"  ]] && echo "✓" || echo "-")"
    done
}

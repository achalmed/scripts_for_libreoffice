#!/usr/bin/env bash
# =============================================================================
# main.sh — pdf-suite: Suite completa de manipulación PDF para Linux
# Autor  : Edison Achalma | UNSCH — Ayacucho, Perú
# Versión: 3.0.0
# GitHub : https://github.com/achalmed
# =============================================================================
# ESTRUCTURA:
#   main.sh            ← estás aquí (orquestador)
#   config.sh          ← variables globales y colores
#   lib/logger.sh      ← logging centralizado
#   lib/validator.sh   ← validación de deps, archivos, argumentos
#   lib/cli.sh         ← parsing CLI + menú interactivo
#   lib/compress.sh    ← comprimir PDFs
#   lib/manipulate.sh  ← merge, split, extract, rotate, reorder, delete
#   lib/convert.sh     ← PDF↔imagen, PDF↔texto, imágenes→PDF
#   lib/metadata.sh    ← leer/escribir metadatos XMP/DocInfo
#   lib/ocr.sh         ← OCR con ocrmypdf + Tesseract
#   lib/protect.sh     ← cifrado, marca de agua, numeración
#   lib/repair.sh      ← reparar, validar, optimizar estructura
# =============================================================================

# Salir inmediatamente en error; tratar variables sin definir como error;
# propagar errores a través de tuberías.
set -euo pipefail

# =============================================================================
# 1. RESOLUCIÓN DE RUTAS
# Resuelve la ruta del script incluso cuando se ejecuta como symlink.
# =============================================================================
_MAIN_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# =============================================================================
# 2. CARGAR MÓDULOS
# El orden importa: config → logger → validator → cli → módulos de dominio
# =============================================================================
# shellcheck source=config.sh
source "${_MAIN_DIR}/config.sh"

# shellcheck source=lib/logger.sh
source "${_MAIN_DIR}/lib/logger.sh"

# shellcheck source=lib/validator.sh
source "${_MAIN_DIR}/lib/validator.sh"

# shellcheck source=lib/cli.sh
source "${_MAIN_DIR}/lib/cli.sh"

# shellcheck source=lib/compress.sh
source "${_MAIN_DIR}/lib/compress.sh"

# shellcheck source=lib/manipulate.sh
source "${_MAIN_DIR}/lib/manipulate.sh"

# shellcheck source=lib/convert.sh
source "${_MAIN_DIR}/lib/convert.sh"

# shellcheck source=lib/metadata.sh
source "${_MAIN_DIR}/lib/metadata.sh"

# shellcheck source=lib/ocr.sh
source "${_MAIN_DIR}/lib/ocr.sh"

# shellcheck source=lib/protect.sh
source "${_MAIN_DIR}/lib/protect.sh"

# shellcheck source=lib/repair.sh
source "${_MAIN_DIR}/lib/repair.sh"

# =============================================================================
# 3. LIMPIEZA AL SALIR
# Garantiza que archivos temporales se eliminen incluso si el script falla.
# =============================================================================
cleanup() {
    [[ -n "${WORK_DIR:-}" && -d "${WORK_DIR}" ]] && rm -rf "${WORK_DIR}"
    [[ -n "${STATS_FILE:-}" && -f "${STATS_FILE}" ]] && rm -f "${STATS_FILE}"
}
trap cleanup EXIT

# Crear directorio de trabajo temporal del script
WORK_DIR="$(mktemp -d /tmp/pdfsuite_work_XXXXXX)"

# =============================================================================
# 4. VALIDACIONES PREVIAS
# =============================================================================
check_required_deps

# =============================================================================
# 5. PARSEO DE ARGUMENTOS GLOBALES
# Extrae flags globales y deja la operación + sus argumentos en REMAINING_ARGS.
# =============================================================================
parse_global_flags "$@"
init_log_file

# =============================================================================
# 6. DESPACHO DE OPERACIONES
# =============================================================================
main() {
    # Sin argumentos → menú interactivo
    if [[ ${#REMAINING_ARGS[@]} -eq 0 ]]; then
        show_interactive_menu
        # INTERACTIVE_OPERATION se establece dentro de show_interactive_menu
        REMAINING_ARGS=("${INTERACTIVE_OPERATION}")
        # Pedir argumentos según la operación seleccionada
        # (cada módulo tiene su propia lógica interactiva cuando no recibe args)
    fi

    local operation="${REMAINING_ARGS[0]}"
    local op_args=("${REMAINING_ARGS[@]:1}")  # argumentos después de la operación

    # Mostrar banner solo en modo interactivo o verbose
    [[ "${VERBOSE:-false}" == "true" || ${#REMAINING_ARGS[@]} -le 1 ]] && \
        show_banner

    log_debug "Operación: ${operation}"
    log_debug "Argumentos: ${op_args[*]:-<ninguno>}"
    [[ "${DRY_RUN:-false}" == "true" ]] && \
        log_warn "MODO DRY-RUN activado: no se escribirá ningún archivo."

    case "$operation" in
        compress)
            run_compress "${op_args[@]}"
            ;;
        merge)
            run_merge "${op_args[@]}"
            ;;
        split)
            run_split "${op_args[@]}"
            ;;
        extract)
            run_extract "${op_args[@]}"
            ;;
        rotate)
            run_rotate "${op_args[@]}"
            ;;
        reorder)
            run_reorder "${op_args[@]}"
            ;;
        delete)
            run_delete "${op_args[@]}"
            ;;
        convert)
            run_convert "${op_args[@]}"
            ;;
        ocr)
            run_ocr "${op_args[@]}"
            ;;
        metadata|meta)
            run_metadata "${op_args[@]}"
            ;;
        protect|encrypt|decrypt)
            # Atajos: 'encrypt' y 'decrypt' invocan run_protect con el flag correcto
            case "$operation" in
                encrypt) run_protect --encrypt "${op_args[@]}" ;;
                decrypt) run_protect --decrypt "${op_args[@]}" ;;
                *)       run_protect "${op_args[@]}" ;;
            esac
            ;;
        watermark)
            run_watermark "${op_args[@]}"
            ;;
        repair)
            run_repair "${op_args[@]}"
            ;;
        validate)
            run_repair --validate "${op_args[@]}"
            ;;
        optimize)
            run_repair --optimize "${op_args[@]}"
            ;;
        info)
            # Alias de metadata sin flags de escritura
            for f in "${op_args[@]}"; do show_pdf_info "$f"; done
            ;;
        test)
            # Prueba todos los métodos de compresión en un PDF
            for f in "${op_args[@]}"; do test_all_methods "$f"; done
            ;;
        deps|check-deps)
            report_optional_deps
            ;;
        -h|--help|help)
            show_help
            ;;
        *)
            log_error "Operación desconocida: '${operation}'"
            echo ""
            show_help
            exit 2
            ;;
    esac
}

main

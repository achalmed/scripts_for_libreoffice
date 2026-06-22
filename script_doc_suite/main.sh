#!/usr/bin/env bash
# main.sh — docflow: Suite unificada de conversión de documentos Office
# Punto de entrada principal. Orquesta todos los módulos en lib/.
#
# Por qué set -euo pipefail aquí y no en los módulos:
#   Los módulos usan "|| return N" para control de flujo deliberado.
#   set -e en un módulo convertiría cada retorno != 0 en abort inesperado.
#   main.sh sí lo necesita para fallar limpio ante errores no manejados.
set -euo pipefail

# ---------------------------------------------------------------------------
# Bootstrapping: cargar módulos desde la ubicación real del script
# (readlink -f resuelve symlinks — fix Bug #5)
# ---------------------------------------------------------------------------
_MAIN_DIR="$(cd "$(dirname "$(readlink -f "${BASH_SOURCE[0]}")")" && pwd)"

# shellcheck source=lib/config.sh
source "${_MAIN_DIR}/lib/config.sh"
# shellcheck source=lib/logger.sh
source "${_MAIN_DIR}/lib/logger.sh"
# shellcheck source=lib/validator.sh
source "${_MAIN_DIR}/lib/validator.sh"
# shellcheck source=lib/cli.sh
source "${_MAIN_DIR}/lib/cli.sh"
# shellcheck source=lib/converter_odf.sh
source "${_MAIN_DIR}/lib/converter_odf.sh"
# shellcheck source=lib/converter_pdf.sh
source "${_MAIN_DIR}/lib/converter_pdf.sh"
# shellcheck source=lib/converter_md.sh
source "${_MAIN_DIR}/lib/converter_md.sh"
# shellcheck source=lib/watcher.sh
source "${_MAIN_DIR}/lib/watcher.sh"
# shellcheck source=lib/reporter.sh
source "${_MAIN_DIR}/lib/reporter.sh"

# ---------------------------------------------------------------------------
# main()
# Orquestador: parsea argumentos y despacha al módulo correcto.
# ---------------------------------------------------------------------------
main() {
    # Parsear argumentos — popula COMMAND y todas las variables globales de configuración
    parse_arguments "$@"
    validate_command

    log_debug "docflow v${VERSION} iniciando"
    log_debug "Comando: ${COMMAND} | Input: ${INPUT_PATH:-<ninguno>} | dry-run: ${DRY_RUN}"

    # Despachar al módulo según el comando
    case "$COMMAND" in
        to-odf)  run_to_odf  ;;
        to-pdf)  run_to_pdf  ;;
        to-md)   run_to_md   ;;
        watch)   run_watch   ;;
        report)  run_report  ;;
    esac

    # Generar reporte automático al final si se especificó formato
    if [[ "${REPORT_FORMAT:-none}" != "none" && "$COMMAND" != "report" && ${#LOG_RESULTS[@]} -gt 0 ]]; then
        run_report
    fi
}

main "$@"

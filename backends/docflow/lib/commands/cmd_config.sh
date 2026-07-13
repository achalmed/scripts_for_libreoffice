#!/usr/bin/env bash
# lib/commands/cmd_config.sh — Comando config: configuración de usuario.

cmd_config() {
    local cfg="${DOCFLOW_CONFIG_FILE:-$DOCFLOW_CONFIG_FILE_DEFAULT}"
    case "${1:-show}" in
        init)
            if [[ -f "$cfg" && "${DOCFLOW_FORCE:-false}" != "true" ]]; then
                log_warn "Ya existe: ${cfg} (usa --force para regenerarla)"
                exit "$EXIT_ERROR"
            fi
            config_write_template "$cfg"
            log_ok "Plantilla creada: ${cfg}"
            ;;
        show)
            if [[ -f "$cfg" ]]; then
                log_info "Configuración: ${cfg}"
                cat -- "$cfg" >&2
            else
                log_info "Sin configuración de usuario (docflow config init para crearla)."
                log_plain "  Valores efectivos actuales:"
                compgen -v DOCFLOW_ | while read -r var; do
                    case "$var" in DOCFLOW_ARGS|DOCFLOW_COMMAND|DOCFLOW_ROOT) continue ;; esac
                    log_plain "    ${var}=${!var}"
                done
            fi
            ;;
        path)
            printf '%s\n' "$cfg"
            ;;
        *)
            log_error "Uso: docflow config [init|show|path]"
            exit "$EXIT_USAGE" ;;
    esac
    exit 0
}

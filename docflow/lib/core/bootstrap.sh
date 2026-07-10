#!/usr/bin/env bash
# lib/core/bootstrap.sh — Cargador de módulos y despachador principal.
# Responsabilidad única: cargar el núcleo en orden, registrar formatos,
# resolver el subcomando y delegar en lib/commands/cmd_<comando>.sh.

# ---------------------------------------------------------------------------
# Carga ordenada del núcleo (el orden importa: constants → utils → logger → …)
# ---------------------------------------------------------------------------
_docflow_load_core() {
    local mod
    for mod in constants utils logger config deps validator registry \
               soffice finder session cache hooks backup progress \
               parallel dispatch cli; do
        # shellcheck disable=SC1090
        source "${DOCFLOW_ROOT}/lib/core/${mod}.sh"
    done
}

# ---------------------------------------------------------------------------
# Carga de motores y formatos. Cada formato se auto-registra en el registry;
# añadir un formato nuevo = añadir un archivo en lib/formats/, nada más.
# ---------------------------------------------------------------------------
_docflow_load_engines() {
    local f
    for f in "${DOCFLOW_ROOT}"/lib/engines/*.sh; do
        # shellcheck disable=SC1090
        source "$f"
    done
}

_docflow_load_formats() {
    local f
    for f in "${DOCFLOW_ROOT}"/lib/formats/*.sh; do
        # shellcheck disable=SC1090
        source "$f"
    done
}

# ---------------------------------------------------------------------------
# docflow_main() — parsea, valida y despacha.
# ---------------------------------------------------------------------------
docflow_main() {
    _docflow_load_core
    _docflow_load_engines
    _docflow_load_formats

    config_load_defaults
    cli_parse_global "$@"

    # cli_parse_global deja: DOCFLOW_COMMAND y DOCFLOW_ARGS (resto de args)
    local cmd="${DOCFLOW_COMMAND:-}"
    case "$cmd" in
        "" ) cli_show_help; return "$EXIT_USAGE" ;;
        help) cli_show_help; return 0 ;;
        version) printf '%s v%s\n' "$DOCFLOW_NAME" "$DOCFLOW_VERSION"; return 0 ;;
    esac

    local cmd_file="${DOCFLOW_ROOT}/lib/commands/cmd_${cmd//-/_}.sh"
    if [[ ! -f "$cmd_file" ]]; then
        log_error "Comando desconocido: '${cmd}'"
        log_plain "  Ejecuta 'docflow help' para ver los comandos disponibles."
        return "$EXIT_USAGE"
    fi

    # shellcheck disable=SC1090
    source "$cmd_file"

    trap 'session_on_interrupt' INT TERM
    "cmd_${cmd//-/_}" "${DOCFLOW_ARGS[@]}"
}

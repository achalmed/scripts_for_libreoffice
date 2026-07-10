#!/usr/bin/env bash
# lib/core/hooks.sh — Hooks de usuario pre/post conversión.
# El comando del hook recibe contexto por variables de entorno:
#   DOCFLOW_INPUT, DOCFLOW_OUTPUT, DOCFLOW_STATUS (solo post), DOCFLOW_TASK.

# hooks_run <pre|post> <input> <output> [status]
hooks_run() {
    local phase="$1" input="$2" output="$3" status="${4:-}"
    local cmd=""
    case "$phase" in
        pre)  cmd="${DOCFLOW_PRE_HOOK:-}" ;;
        post) cmd="${DOCFLOW_POST_HOOK:-}" ;;
    esac
    [[ -z "$cmd" ]] && return 0
    [[ "${DOCFLOW_DRY_RUN:-false}" == "true" ]] && return 0

    log_trace "hook ${phase}: ${cmd}"
    if ! DOCFLOW_INPUT="$input" DOCFLOW_OUTPUT="$output" \
         DOCFLOW_STATUS="$status" DOCFLOW_TASK="${DOCFLOW_TASK:-}" \
         bash -c "$cmd" >/dev/null 2>&1; then
        log_warn "Hook ${phase} devolvió error para: $(basename -- "$input")"
    fi
    return 0
}

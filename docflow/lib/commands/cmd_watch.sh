#!/usr/bin/env bash
# lib/commands/cmd_watch.sh — Comando watch: vigilar y convertir al vuelo.

cmd_watch() {
    local task="${DOCFLOW_WATCH_TASK:-pdf}" dir=""
    local -a rest=()
    while [[ $# -gt 0 ]]; do
        case "$1" in
            --to) task="$2"; shift 2 ;;
            *) rest+=("$1"); shift ;;
        esac
    done
    dir="${rest[0]:-}"

    case "$task" in md|pdf|odf|office) ;; *)
        log_error "Destino inválido para watch: '$task' (usa md|pdf|odf|office)"
        exit "$EXIT_USAGE" ;;
    esac
    [[ -d "$dir" ]] || { log_error "Se requiere un directorio a vigilar."; cli_help_watch; exit "$EXIT_USAGE"; }

    watch_run "$task" "$dir"
}

cli_help_watch() {
    cat >&2 <<EOF
${C_BOLD}docflow watch${C_RESET} — Vigila un directorio y convierte automáticamente.

USO: docflow watch [--to md|pdf|odf|office] [opciones] <directorio>

  Usa inotify si está disponible (detección instantánea, CPU cero);
  si no, polling cada --interval segundos.

OPCIONES ESPECÍFICAS:
  --to <destino>      Tipo de conversión (defecto: pdf)
  --interval <seg>    Intervalo de polling (defecto: 30)
  --post-hook <cmd>   Comando tras cada conversión (recibe \$DOCFLOW_OUTPUT)

EJEMPLOS:
  docflow watch ~/Descargas
  docflow watch --to md -f docx,odt ~/Documentos/notas/
  docflow watch ~/Descargas --post-hook 'notify-send docflow "\$DOCFLOW_OUTPUT"'
EOF
}

#!/usr/bin/env bash
# lib/engines/watch.sh — Watch Engine.
# Vigila un directorio y convierte archivos nuevos/modificados al vuelo.
# inotify (eventos del kernel, CPU cero en espera) con fallback a polling.

# watch_run <task> <directorio>
watch_run() {
    local task="$1" dir="$2"
    local exts_csv
    exts_csv="$(_dispatch_exts_for_task "$task")"
    export DOCFLOW_TASK="$task"
    session_init "watch-${task}"

    log_header "Vigilando ${dir} → ${task} (Ctrl-C para salir)"
    log_info "Formatos: ${exts_csv}"

    if deps_has inotifywait; then
        _watch_inotify "$task" "$dir" "$exts_csv"
    else
        log_warn "inotify-tools no instalado; usando polling cada ${DOCFLOW_WATCH_INTERVAL}s"
        log_plain "  Para vigilancia instantánea: $(deps_install_hint inotifywait)"
        _watch_polling "$task" "$dir" "$exts_csv"
    fi
}

# _watch_matches <archivo> <exts_csv> → 0 si la extensión está en la lista
_watch_matches() {
    local ext exts
    ext="$(util_ext "$1")"
    IFS=',' read -ra exts <<< "$2"
    util_in_list "$ext" "${exts[@]}"
}

_watch_handle() {
    local task="$1" file="$2"
    # Espera de asentamiento: sincronizadores (Nextcloud, etc.) pueden emitir
    # el evento antes de terminar de escribir el archivo.
    sleep "${DOCFLOW_WATCH_SETTLE:-1}"
    [[ -f "$file" ]] || return 0
    log_info "Detectado: $(basename -- "$file")"
    dispatch_convert_one "$task" "$file"
}

_watch_inotify() {
    local task="$1" dir="$2" exts_csv="$3" file
    inotifywait -m -r -q -e close_write -e moved_to --format '%w%f' -- "$dir" \
    | while IFS= read -r file; do
        _watch_matches "$file" "$exts_csv" || continue
        case "$(basename -- "$file")" in '~$'*|'.~lock.'*) continue ;; esac
        _watch_handle "$task" "$file"
    done
}

_watch_polling() {
    local task="$1" dir="$2" exts_csv="$3"
    local snapshot new_snapshot file
    snapshot="$(mktemp "${TMPDIR:-/tmp}/docflow_watch_XXXXXX")"
    new_snapshot="$(mktemp "${TMPDIR:-/tmp}/docflow_watch_XXXXXX")"
    trap 'rm -f -- "$snapshot" "$new_snapshot"' RETURN

    _watch_snapshot "$dir" > "$snapshot"
    while true; do
        sleep "${DOCFLOW_WATCH_INTERVAL:-30}"
        _watch_snapshot "$dir" > "$new_snapshot"
        # Líneas nuevas o cambiadas (ruta + mtime) respecto al snapshot anterior
        while IFS= read -r file; do
            file="${file#* }"
            _watch_matches "$file" "$exts_csv" || continue
            _watch_handle "$task" "$file"
        done < <(comm -13 "$snapshot" "$new_snapshot")
        cp -- "$new_snapshot" "$snapshot"
    done
}

# _watch_snapshot <dir> → "mtime ruta" ordenado (para comm)
_watch_snapshot() {
    find "$1" -type f -printf '%T@ %p\n' 2>/dev/null | sort
}

#!/usr/bin/env bash
# lib/watcher.sh — Vigilancia de carpeta y conversión automática
# Por qué existe: detectar cambios con inotifywait es O(1) en eventos,
# mucho más eficiente que un loop con find cada N segundos.
# Cuando inotify no está disponible, cae de forma elegante a polling.

# ---------------------------------------------------------------------------
# _detect_watch_mode()
# Detecta si inotifywait está disponible.
# Imprime: "inotify" o "poll"
# ---------------------------------------------------------------------------
_detect_watch_mode() {
    if command -v inotifywait &>/dev/null; then
        echo "inotify"
    else
        echo "poll"
    fi
}

# ---------------------------------------------------------------------------
# _dispatch_conversion()
# Despacha un archivo recién detectado al conversor correcto según su extensión.
# Argumentos: $1=ruta_del_archivo
# ---------------------------------------------------------------------------
_dispatch_conversion() {
    local file="$1"
    local ext="${file##*.}"
    ext="${ext,,}"

    log_info "Archivo detectado: $(basename "$file")"

    # Guardar INPUT_PATH actual y restaurar después
    local prev_input="$INPUT_PATH"
    INPUT_PATH="$file"

    case "$ext" in
        docx|doc|odt)
            # Decidir target según COMMAND original del watch
            case "${WATCH_TARGET_CMD:-to-pdf}" in
                to-odf) run_to_odf ;;
                to-md)  run_to_md ;;
                *)      run_to_pdf ;;
            esac
            ;;
        pptx|ppt|odp|ppsx|pps|pptm)
            case "${WATCH_TARGET_CMD:-to-pdf}" in
                to-odf) run_to_odf ;;
                to-md)  run_to_md ;;
                *)      run_to_pdf ;;
            esac
            ;;
        xlsx|xls|ods|xlsm|xltx)
            case "${WATCH_TARGET_CMD:-to-pdf}" in
                to-odf) run_to_odf ;;
                *)      run_to_pdf ;;
            esac
            ;;
        *)
            log_debug "Extensión no manejada: .${ext} — ignorando"
            ;;
    esac

    INPUT_PATH="$prev_input"

    # Ejecutar comando post-conversión si se especificó
    if [[ -n "${WATCH_CMD:-}" ]]; then
        log_debug "Ejecutando post-cmd: $WATCH_CMD"
        eval "$WATCH_CMD" || log_warn "Post-cmd falló: $WATCH_CMD"
    fi
}

# ---------------------------------------------------------------------------
# _build_inotify_extensions()
# Construye la lista de extensiones para el filtro de inotifywait.
# ---------------------------------------------------------------------------
_build_inotify_extensions() {
    local -a all_exts=(
        docx doc odt dotx
        xlsx xls ods xlsm xltx
        pptx ppt odp pptm ppsx pps potx
    )
    if [[ -n "${FORMATS_FILTER:-}" ]]; then
        IFS=',' read -ra all_exts <<< "$FORMATS_FILTER"
    fi

    # inotifywait usa expresiones regulares básicas
    local pattern
    pattern="$(printf '\.%s$|' "${all_exts[@]}")"
    echo "${pattern%|}"  # quitar el último |
}

# ---------------------------------------------------------------------------
# _watch_inotify()
# Usa inotifywait para vigilancia eficiente basada en eventos del kernel.
# Solo dispara cuando un archivo se cierra tras escritura (CLOSE_WRITE).
# ---------------------------------------------------------------------------
_watch_inotify() {
    local watch_dir="$1"
    local ext_pattern
    ext_pattern="$(_build_inotify_extensions)"

    log_info "Modo inotify activo — vigilando: $watch_dir"
    log_info "Presiona Ctrl+C para detener."
    echo ""

    # --monitor: no terminar después del primer evento
    # --recursive: subdirectorios incluidos
    # --event CLOSE_WRITE: solo cuando el archivo se ha escrito y cerrado
    #   (evita procesar archivos a medias mientras se guardan)
    inotifywait \
        --monitor \
        --recursive \
        --event CLOSE_WRITE \
        --format '%w%f' \
        --include "$ext_pattern" \
        "$watch_dir" 2>/dev/null \
    | while IFS= read -r new_file; do
        # Pequeña pausa para que el proceso que escribió termine de soltar el archivo
        sleep 0.5
        _dispatch_conversion "$new_file"
    done
}

# ---------------------------------------------------------------------------
# _watch_poll()
# Fallback: polling cada WATCH_INTERVAL segundos.
# Mantiene un snapshot de los archivos vistos para detectar nuevos.
# ---------------------------------------------------------------------------
_watch_poll() {
    local watch_dir="$1"
    local interval="${WATCH_INTERVAL:-30}"

    log_warn "inotifywait no disponible — usando polling cada ${interval}s"
    log_warn "Para vigilancia eficiente: sudo pacman -S inotify-tools"
    log_info "Presiona Ctrl+C para detener."
    echo ""

    local snapshot_file
    snapshot_file="$(mktemp /tmp/docflow_watch_snap_XXXXXX)"
    trap "rm -f $snapshot_file" RETURN

    # Snapshot inicial: todos los archivos existentes
    find "$watch_dir" -type f \( \
        -iname "*.docx" -o -iname "*.doc" -o -iname "*.odt" -o \
        -iname "*.pptx" -o -iname "*.ppt" -o -iname "*.odp" -o \
        -iname "*.xlsx" -o -iname "*.xls" -o -iname "*.ods" \
    \) -print0 2>/dev/null | sort -z > "$snapshot_file"

    while true; do
        sleep "$interval"

        local new_snapshot
        new_snapshot="$(mktemp /tmp/docflow_watch_new_XXXXXX)"

        find "$watch_dir" -type f \( \
            -iname "*.docx" -o -iname "*.doc" -o -iname "*.odt" -o \
            -iname "*.pptx" -o -iname "*.ppt" -o -iname "*.odp" -o \
            -iname "*.xlsx" -o -iname "*.xls" -o -iname "*.ods" \
        \) -print0 2>/dev/null | sort -z > "$new_snapshot"

        # Detectar archivos nuevos comparando con snapshot anterior
        while IFS= read -r -d '' new_file; do
            if ! grep -qF "$new_file" "$snapshot_file" 2>/dev/null; then
                _dispatch_conversion "$new_file"
            fi
        done < "$new_snapshot"

        mv "$new_snapshot" "$snapshot_file"
    done
}

# ---------------------------------------------------------------------------
# run_watch()
# Punto de entrada público del comando watch.
# Argumentos globales: INPUT_PATH, FORMATS_FILTER, WATCH_INTERVAL
# ---------------------------------------------------------------------------
run_watch() {
    log_header "MODO WATCH — Conversión automática"

    if [[ -z "${INPUT_PATH:-}" ]]; then
        log_error "Se requiere un directorio para vigilar."
        echo "  Usa: docflow watch -i <directorio> [--formats docx,pptx]"
        exit 3
    fi

    validate_path_exists "$INPUT_PATH" dir || exit 3

    # Guardar el comando de conversión objetivo (to-pdf por defecto en watch)
    WATCH_TARGET_CMD="${WATCH_CMD_TARGET:-to-pdf}"
    log_info "Conversión automática → ${WATCH_TARGET_CMD}"

    local watch_mode
    watch_mode="$(_detect_watch_mode)"

    if [[ "$watch_mode" == "inotify" ]]; then
        _watch_inotify "$INPUT_PATH"
    else
        _watch_poll "$INPUT_PATH"
    fi
}

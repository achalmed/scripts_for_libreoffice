#!/usr/bin/env bash
# lib/core/dispatch.sh — Pipeline común de todo lote de conversión.
# Los comandos to-md/to-pdf/to-odf/to-office son envoltorios finos sobre:
#   dispatch_run <tarea> <rutas...>
# donde <tarea> ∈ md | pdf | odf | office. Cada motor implementa el contrato:
#   engine_<tarea>_output_ext <ext_entrada>  → extensión de salida ('' = n/a)
#   engine_<tarea>_convert    <in> <out>     → realiza la conversión
# Este módulo aporta todo lo transversal: descubrimiento, caché, hooks,
# overwrite/resume, verificación, sesión, progreso, paralelo y originales.

# dispatch_run <task> <rutas...>
dispatch_run() {
    local task="$1"
    shift
    export DOCFLOW_TASK="$task"

    local exts_csv
    exts_csv="$(_dispatch_exts_for_task "$task")"

    local -a targets=()
    finder_resolve_targets targets "$exts_csv" "$@" || return "$EXIT_BAD_INPUT"

    if [[ ${#targets[@]} -eq 0 ]]; then
        log_warn "No se encontraron archivos convertibles en: $*"
        return "$EXIT_NO_FILES"
    fi

    # Raíz del lote para preservar estructura relativa con --output-dir
    if [[ -d "${1:-}" ]]; then
        DOCFLOW_BATCH_ROOT="$(readlink -f -- "$1")"
    else
        DOCFLOW_BATCH_ROOT="$(dirname -- "$(readlink -f -- "$1")")"
    fi
    export DOCFLOW_BATCH_ROOT

    session_init "$task"
    log_info "Archivos a procesar: ${#targets[@]}"

    if [[ "${DOCFLOW_DRY_RUN:-false}" == "true" ]]; then
        _dispatch_sequential "$task" "${targets[@]}"
    elif (( ${#targets[@]} > 2 && DOCFLOW_JOBS > 1 )); then
        local list
        list="$(mktemp "${TMPDIR:-/tmp}/docflow_batch_XXXXXX")"
        printf '%s\0' "${targets[@]}" > "$list"
        parallel_run "$task" "${#targets[@]}" "$list"
        rm -f -- "$list"
    else
        _dispatch_sequential "$task" "${targets[@]}"
    fi

    backup_handle_originals
    _dispatch_auto_report
    session_prune

    session_summary
}

_dispatch_sequential() {
    local task="$1"
    shift
    local start_ms i=0 total=$#
    start_ms="$(util_epoch_ms)"
    local file
    for file in "$@"; do
        dispatch_convert_one "$task" "$file"
        i=$(( i + 1 ))
        progress_enabled && progress_draw "$i" "$total" "$start_ms"
    done
    return 0
}

# _dispatch_exts_for_task <task> → extensiones aplicables (CSV), respetando -f
_dispatch_exts_for_task() {
    local task="$1" field
    case "$task" in
        md) field="md" ;; pdf) field="pdf" ;; odf) field="odf" ;; office) field="ms" ;;
        *) field="all" ;;
    esac
    if [[ -n "${DOCFLOW_FORMATS_FILTER:-}" ]]; then
        printf '%s' "$DOCFLOW_FORMATS_FILTER"
    else
        registry_exts_for "$field" | paste -sd,
    fi
}

# _dispatch_output_path <task> <input> <out_ext> → ruta de salida calculada
_dispatch_output_path() {
    local task="$1" input="$2" out_ext="$3"
    local base out_dir
    base="$(util_basename_noext "$input")"

    if [[ -n "${DOCFLOW_OUTPUT_DIR:-}" ]]; then
        # Con -o: preservar la estructura relativa a la raíz del lote
        local rel
        rel="$(util_relpath "$(dirname -- "$(readlink -f -- "$input")")" "$DOCFLOW_BATCH_ROOT")"
        [[ "$rel" == "." ]] && out_dir="$DOCFLOW_OUTPUT_DIR" || out_dir="${DOCFLOW_OUTPUT_DIR}/${rel}"
    else
        out_dir="$(dirname -- "$input")"
    fi
    printf '%s/%s.%s' "$out_dir" "$base" "$out_ext"
}

# dispatch_convert_one <task> <input>
# Conversión de UN archivo con todo el ciclo de vida. Registra el resultado
# en la sesión; siempre devuelve 0 (el lote no se detiene por un archivo).
dispatch_convert_one() {
    local task="$1" input="$2"
    local ext out_ext output
    ext="$(util_ext "$input")"

    if ! registry_supported "$ext"; then
        log_skip "formato no soportado: $(basename -- "$input")"
        session_record skip "$task" "$input" "" 0 0 0 "formato no soportado"
        return 0
    fi

    out_ext="$("engine_${task}_output_ext" "$ext")"
    if [[ -z "$out_ext" ]]; then
        log_skip "sin ruta ${task} para .${ext}: $(basename -- "$input")"
        session_record skip "$task" "$input" "" 0 0 0 "sin conversión ${task} para .${ext}"
        return 0
    fi

    output="$(_dispatch_output_path "$task" "$input" "$out_ext")"

    # Mismo archivo de entrada y salida (p. ej. pdf→pdf) → nada que hacer
    if [[ "$(readlink -f -- "$input")" == "$(readlink -f -- "$output" 2>/dev/null)" ]]; then
        session_record skip "$task" "$input" "$output" 0 0 0 "entrada = salida"
        return 0
    fi

    # Overwrite / resume: si la salida existe y no se pide sobreescribir
    if [[ -s "$output" && "${DOCFLOW_OVERWRITE:-false}" != "true" ]]; then
        log_skip "ya existe: $(basename -- "$output")"
        session_record skip "$task" "$input" "$output" 0 \
            "$(util_file_size "$input")" "$(util_file_size "$output")" "salida ya existe"
        return 0
    fi

    # Caché por contenido: mismo archivo + mismas opciones → no reconvertir
    if [[ -s "$output" ]] && cache_check "$input" "$task" "$output"; then
        log_skip "sin cambios (caché): $(basename -- "$input")"
        session_record cached "$task" "$input" "$output" 0 \
            "$(util_file_size "$input")" "$(util_file_size "$output")" ""
        return 0
    fi

    # Documentos protegidos por contraseña: reporte específico, nunca "fallo mudo"
    if validate_protected "$input"; then
        log_fail "protegido por contraseña o corrupto: $(basename -- "$input")"
        session_record protected "$task" "$input" "$output" 0 \
            "$(util_file_size "$input")" 0 "documento protegido o corrupto"
        return 0
    fi

    if [[ "${DOCFLOW_DRY_RUN:-false}" == "true" ]]; then
        log_info "[dry-run] $(basename -- "$input") → ${output}"
        session_record dry-run "$task" "$input" "$output" 0 \
            "$(util_file_size "$input")" 0 ""
        return 0
    fi

    validate_writable_dir "$(dirname -- "$output")" || {
        session_record fail "$task" "$input" "$output" 0 0 0 "directorio de salida no escribible"
        return 0
    }

    hooks_run pre "$input" "$output"

    # Si la salida no existía, un fallo no debe dejar restos parciales que
    # luego bloqueen el reintento con "ya existe".
    local output_preexisted=false
    [[ -f "$output" ]] && output_preexisted=true

    local t0 t1 rc
    t0="$(util_epoch_ms)"
    "engine_${task}_convert" "$input" "$output"
    rc=$?
    t1="$(util_epoch_ms)"

    if (( rc != 0 )); then
        log_fail "$(basename -- "$input")"
        session_record fail "$task" "$input" "$output" $((t1 - t0)) \
            "$(util_file_size "$input")" 0 "el motor de conversión devolvió error"
        _dispatch_cleanup_partial "$task" "$output" "$output_preexisted"
        hooks_run post "$input" "$output" "fail"
        return 0
    fi

    # Verificación post-conversión: si falla, se marca fail (el original
    # NUNCA se contará como convertido con éxito).
    if [[ "${DOCFLOW_VERIFY:-true}" == "true" ]] && ! validation_verify "$task" "$output"; then
        log_fail "verificación fallida: $(basename -- "$output")"
        session_record fail "$task" "$input" "$output" $((t1 - t0)) \
            "$(util_file_size "$input")" "$(util_file_size "$output")" "verificación post-conversión fallida"
        _dispatch_cleanup_partial "$task" "$output" "$output_preexisted"
        hooks_run post "$input" "$output" "fail"
        return 0
    fi

    cache_store "$input" "$task" "$output"
    log_ok "$(basename -- "$input") → $(basename -- "$output") ($(( t1 - t0 )) ms)"
    session_record ok "$task" "$input" "$output" $((t1 - t0)) \
        "$(util_file_size "$input")" "$(util_file_size "$output")" ""
    hooks_run post "$input" "$output" "ok"
    return 0
}

# _dispatch_cleanup_partial <task> <output> <preexistía>
# Elimina la salida parcial de una conversión fallida (solo si docflow la creó).
_dispatch_cleanup_partial() {
    local task="$1" output="$2" preexisted="$3"
    [[ "$preexisted" == "true" ]] && return 0
    rm -f -- "$output"
    [[ "$task" == "md" ]] && rm -rf -- "${output%.md}_files"
    return 0
}

# _dispatch_auto_report — genera reportes si se pidieron con --report
_dispatch_auto_report() {
    [[ -z "${DOCFLOW_REPORT_FORMATS:-}" ]] && return 0
    reporting_generate "$DOCFLOW_SESSION_DIR" "$DOCFLOW_REPORT_FORMATS" "${DOCFLOW_REPORT_OUT:-}"
}

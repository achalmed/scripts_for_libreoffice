#!/usr/bin/env bash
# lib/core/parallel.sh — Ejecución paralela de conversiones.
# Diseño: en lugar de exportar decenas de funciones a subshells (frágil),
# cada worker re-invoca `bin/docflow __worker <tarea>` — un subcomando
# interno que re-carga el entorno y convierte UN archivo. El estado se
# comparte vía variables DOCFLOW_* exportadas y el results.tsv de la sesión.

# parallel_export_env — exporta toda la configuración efectiva a los workers
parallel_export_env() {
    local var
    while IFS= read -r var; do
        # shellcheck disable=SC2163
        export "$var"
    done < <(compgen -v DOCFLOW_)
}

# parallel_run <task> <total> <archivo_nul_list>
# Lanza los workers con xargs -P y monitoriza el progreso desde results.tsv.
parallel_run() {
    local task="$1" total="$2" file_list="$3"
    local jobs="${DOCFLOW_JOBS:-2}"
    local start_ms results="${DOCFLOW_SESSION_DIR}/results.tsv"
    start_ms="$(util_epoch_ms)"

    parallel_export_env
    log_info "Procesando en paralelo con ${jobs} workers..."

    # xargs en segundo plano; el proceso principal dibuja el progreso.
    xargs -0 -P "$jobs" -n 1 \
        "${DOCFLOW_ROOT}/bin/docflow" __worker "$task" \
        < "$file_list" &
    local xargs_pid=$!

    if progress_enabled; then
        local done_count=0
        while kill -0 "$xargs_pid" 2>/dev/null; do
            done_count="$(wc -l < "$results" 2>/dev/null || echo 0)"
            progress_draw "$done_count" "$total" "$start_ms"
            sleep 0.3
        done
        progress_draw "$(wc -l < "$results")" "$total" "$start_ms"
    fi

    wait "$xargs_pid"
    return 0
}

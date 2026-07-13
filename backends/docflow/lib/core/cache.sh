#!/usr/bin/env bash
# lib/core/cache.sh — Caché de conversiones por contenido (SHA-256).
# Evita reconvertir archivos que no han cambiado: la clave es
# hash(contenido) + tarea + opciones relevantes; el valor, la ruta de salida.
# Índice en ~/.cache/docflow/index.tsv (append-only; la última entrada gana).

_CACHE_INDEX="${DOCFLOW_CACHE_DIR}/index.tsv"

# _cache_key <input> <task> → clave estable de la conversión
_cache_key() {
    local input="$1" task="$2" content_hash
    content_hash="$(util_sha256 "$input")" || return 1
    # Las opciones que cambian el resultado forman parte de la clave:
    printf '%s' "${content_hash}:${task}:${DOCFLOW_MD_FLAVOR:-}:${DOCFLOW_IMG_FORMAT:-}:${DOCFLOW_PDF_ENGINE:-}" \
        | sha256sum | cut -d' ' -f1
}

# cache_check <input> <task> <output_esperado> → 0 si hay hit válido
cache_check() {
    [[ "${DOCFLOW_CACHE:-true}" != "true" ]] && return 1
    [[ -f "$_CACHE_INDEX" ]] || return 1
    local key cached_output
    key="$(_cache_key "$1" "$2")" || return 1
    cached_output="$(awk -F'\t' -v k="$key" '$1 == k { out = $2 } END { print out }' "$_CACHE_INDEX")"
    # Hit solo si la salida registrada coincide con la esperada y aún existe
    [[ -n "$cached_output" && "$cached_output" == "$3" && -s "$cached_output" ]]
}

# cache_store <input> <task> <output>
cache_store() {
    [[ "${DOCFLOW_CACHE:-true}" != "true" ]] && return 0
    mkdir -p "$DOCFLOW_CACHE_DIR"
    local key
    key="$(_cache_key "$1" "$2")" || return 0
    printf '%s\t%s\t%s\n' "$key" "$3" "$(date +%s)" >> "$_CACHE_INDEX"
}

# cache_stats — información legible del estado del caché
cache_stats() {
    if [[ ! -f "$_CACHE_INDEX" ]]; then
        log_info "Caché vacío."
        return 0
    fi
    local entries size
    entries="$(wc -l < "$_CACHE_INDEX")"
    size="$(util_human_size "$(util_file_size "$_CACHE_INDEX")")"
    log_info "Entradas de caché: ${entries} (índice: ${size})"
    log_plain "  Ubicación: ${_CACHE_INDEX}"
}

# cache_clear — elimina el índice completo
cache_clear() {
    rm -f "$_CACHE_INDEX"
    log_info "Caché eliminado."
}

# cache_prune — descarta entradas cuya salida ya no existe y deduplica
cache_prune() {
    [[ -f "$_CACHE_INDEX" ]] || return 0
    local tmp
    tmp="$(mktemp)"
    # tac + awk '!seen': conserva solo la entrada más reciente de cada clave
    tac "$_CACHE_INDEX" | awk -F'\t' '!seen[$1]++' | while IFS=$'\t' read -r key out ts; do
        [[ -s "$out" ]] && printf '%s\t%s\t%s\n' "$key" "$out" "$ts"
    done > "$tmp"
    mv "$tmp" "$_CACHE_INDEX"
    log_info "Caché depurado: $(wc -l < "$_CACHE_INDEX") entradas vigentes."
}

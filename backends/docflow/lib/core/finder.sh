#!/usr/bin/env bash
# lib/core/finder.sh — Descubrimiento de archivos a procesar.
# Soporta: recursión, filtro por extensión, exclusión por nombre/ruta/glob,
# exclusión por regex, inclusión por glob y rutas con cualquier carácter
# (todo el pipeline usa separadores NUL).

# finder_collect <array_nameref> <raíz> <ext1,ext2,...>
# Aplica: DOCFLOW_EXCLUDES (nombres/globs/rutas), DOCFLOW_EXCLUDE_REGEX,
#         DOCFLOW_INCLUDE_GLOB.
finder_collect() {
    local -n _out="$1"
    local root="$2" exts_csv="$3"
    local -a exts name_args=() prune_args=()
    IFS=',' read -ra exts <<< "$exts_csv"

    local ext
    for ext in "${exts[@]}"; do
        [[ -z "$ext" ]] && continue
        name_args+=(-iname "*.${ext}" -o)
    done
    [[ ${#name_args[@]} -eq 0 ]] && return 0
    unset 'name_args[-1]'

    # Exclusiones: ruta absoluta → -path; con comodín o nombre → -name
    local -a excludes=()
    read -ra excludes <<< "${DOCFLOW_EXCLUDES:-}"
    local excl
    if [[ ${#excludes[@]} -gt 0 ]]; then
        prune_args+=(\()
        for excl in "${excludes[@]}"; do
            if [[ "$excl" == /* ]]; then
                prune_args+=(-path "$excl" -o -path "${excl%/}/*" -o)
            else
                prune_args+=(-name "$excl" -o)
            fi
        done
        unset 'prune_args[-1]'
        prune_args+=(\) -prune -o)
    fi

    local -a find_cmd=(find "$root")
    [[ ${#prune_args[@]} -gt 0 ]] && find_cmd+=("${prune_args[@]}")
    find_cmd+=(-type f \( "${name_args[@]}" \))
    [[ -n "${DOCFLOW_INCLUDE_GLOB:-}" ]] && find_cmd+=(-name "$DOCFLOW_INCLUDE_GLOB")
    [[ -n "${DOCFLOW_EXCLUDE_REGEX:-}" ]] && find_cmd+=(! -regex ".*${DOCFLOW_EXCLUDE_REGEX}.*")
    find_cmd+=(-print0)

    local f
    while IFS= read -r -d '' f; do
        # Ignorar archivos de bloqueo de Office/LibreOffice abiertos (~$doc.docx, .~lock.*)
        case "$(basename -- "$f")" in
            '~$'*|'.~lock.'*) continue ;;
        esac
        _out+=("$f")
    done < <("${find_cmd[@]}" 2>/dev/null | sort -z)
}

# finder_resolve_targets <array_nameref> <exts_csv> <rutas...>
# Acepta mezcla de archivos y directorios; valida y expande.
finder_resolve_targets() {
    local -n _tgt="$1"
    local exts_csv="$2"
    shift 2
    local path
    for path in "$@"; do
        validate_input_path "$path" || return 1
        if [[ -f "$path" ]]; then
            _tgt+=("$path")
        else
            finder_collect _tgt "$path" "$exts_csv"
        fi
    done
    return 0
}

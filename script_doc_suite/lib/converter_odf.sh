#!/usr/bin/env bash
# lib/converter_odf.sh — Motor de conversión MS Office → OpenDocument
# Por qué existe: aislar la lógica de conversión ODF permite probarla
# de forma independiente y reutilizarla desde watch y desde CLI.

# Archivo temporal para registrar conversiones exitosas (evita el bug del
# array asociativo que no se exporta a subshells)
_ODF_SUCCESS_LOG=""

# ---------------------------------------------------------------------------
# _init_odf_session()
# Prepara el archivo temporal de registro para la sesión actual.
# ---------------------------------------------------------------------------
_init_odf_session() {
    _ODF_SUCCESS_LOG="$(mktemp /tmp/docflow_odf_XXXXXX)"
    log_debug "Registro de sesión ODF: $_ODF_SUCCESS_LOG"
}

# ---------------------------------------------------------------------------
# _cleanup_odf_session()
# Elimina el archivo temporal al terminar.
# ---------------------------------------------------------------------------
_cleanup_odf_session() {
    [[ -n "$_ODF_SUCCESS_LOG" && -f "$_ODF_SUCCESS_LOG" ]] && rm -f "$_ODF_SUCCESS_LOG"
}

# ---------------------------------------------------------------------------
# _get_odf_output_ext()
# Dado el formato de entrada, devuelve la extensión ODF de salida.
# Imprime la extensión o vacío si no hay mapeo.
# Argumento: $1=extensión de entrada (sin punto, lowercase)
# ---------------------------------------------------------------------------
_get_odf_output_ext() {
    local in_ext="${1,,}"
    for entry in "${ODF_CONVERSION_MAP[@]}"; do
        local src="${entry%%:*}"
        local rest="${entry#*:}"
        local dst="${rest%%:*}"
        if [[ "$src" == "$in_ext" ]]; then
            echo "$dst"
            return 0
        fi
    done
    return 1
}

# ---------------------------------------------------------------------------
# _convert_single_to_odf()
# Convierte un único archivo a su equivalente ODF.
# Argumentos: $1=ruta_absoluta_archivo, $2=directorio_salida (opcional)
# Retorna: 0 éxito, 1 error, 2 omitido (ya existe)
# ---------------------------------------------------------------------------
_convert_single_to_odf() {
    local input_file="$1"
    local out_dir="${2:-$(dirname "$input_file")}"
    local ext="${input_file##*.}"
    ext="${ext,,}"

    local out_ext
    out_ext="$(_get_odf_output_ext "$ext")" || {
        log_warn "Sin mapeo ODF para '.${ext}': $(basename "$input_file")"
        return 1
    }

    local out_file="${out_dir}/$(basename "${input_file%.*}").${out_ext}"

    # Omitir si ya existe y no se pidió sobreescribir
    if [[ -f "$out_file" && "${OVERWRITE:-false}" != "true" ]]; then
        log_detail "$(basename "$input_file") → ya existe, omitido" "⊘"
        record_result "$input_file" "$out_file" "skip" "to-odf"
        echo "$input_file" >> "$_ODF_SUCCESS_LOG"
        return 2
    fi

    log_detail "$(basename "$input_file") → $(basename "$out_file")"

    if [[ "${DRY_RUN:-false}" == "true" ]]; then
        log_debug "[dry-run] soffice --convert-to $out_ext --outdir $out_dir $input_file"
        record_result "$input_file" "$out_file" "dry-run" "to-odf"
        return 0
    fi

    if soffice --headless --convert-to "$out_ext" --outdir "$out_dir" "$input_file" &>/dev/null; then
        log_detail "✓ $(basename "$out_file")" "✓"
        record_result "$input_file" "$out_file" "ok" "to-odf"
        echo "$input_file" >> "$_ODF_SUCCESS_LOG"
        return 0
    else
        log_error "Falló la conversión: $(basename "$input_file")"
        record_result "$input_file" "$out_file" "fail" "to-odf"
        return 1
    fi
}

# ---------------------------------------------------------------------------
# _handle_originals_odf()
# Gestiona los archivos originales después de la conversión según DELETE_MODE.
# Lee la lista de convertidos exitosos del archivo temporal de sesión.
# ---------------------------------------------------------------------------
_handle_originals_odf() {
    [[ ! -f "$_ODF_SUCCESS_LOG" ]] && return 0
    local converted_count
    converted_count="$(wc -l < "$_ODF_SUCCESS_LOG")"
    [[ "$converted_count" -eq 0 ]] && return 0

    if [[ "${BACKUP_ORIGINALS:-false}" == "true" ]]; then
        log_header "MOVIENDO ORIGINALES A _backup/"
        while IFS= read -r file; do
            local backup_dir
            backup_dir="$(dirname "$file")/${BACKUP_SUFFIX}"
            mkdir -p "$backup_dir"
            if [[ "${DRY_RUN:-false}" == "true" ]]; then
                log_detail "[dry-run] mv $(basename "$file") → $backup_dir/" "↦"
            else
                mv "$file" "$backup_dir/" && log_detail "↦ $(basename "$file") → _backup/" "↦"
            fi
        done < "$_ODF_SUCCESS_LOG"
        return 0
    fi

    if [[ "${DELETE_ORIGINALS:-false}" == "true" ]]; then
        log_header "ELIMINACIÓN DE ORIGINALES"
        if confirm_action "¿Eliminar $converted_count archivo(s) original(es) convertidos?"; then
            while IFS= read -r file; do
                if [[ "${DRY_RUN:-false}" == "true" ]]; then
                    log_detail "[dry-run] rm $(basename "$file")" "🗑"
                else
                    rm "$file" && log_detail "🗑 Eliminado: $(basename "$file")" "🗑"
                fi
            done < "$_ODF_SUCCESS_LOG"
        else
            log_warn "Originales conservados (decisión del usuario)."
        fi
    fi
}

# ---------------------------------------------------------------------------
# run_to_odf()
# Punto de entrada público del comando to-odf.
# Procesa INPUT_PATH como archivo único o directorio completo.
# ---------------------------------------------------------------------------
run_to_odf() {
    log_header "CONVERSIÓN: MS Office → OpenDocument Format"
    validate_dependencies soffice || exit 5

    _init_odf_session
    trap _cleanup_odf_session EXIT

    local -a targets=()

    # Determinar si se procesará un archivo único o un directorio
    if [[ -f "${INPUT_PATH:-}" ]]; then
        targets+=("$INPUT_PATH")
    elif [[ -d "${INPUT_PATH:-}" ]]; then
        _collect_odf_targets targets "$INPUT_PATH"
    else
        log_error "Ruta inválida o no especificada: '${INPUT_PATH:-}'"
        echo "  Usa: docflow to-odf -i <archivo_o_directorio>"
        exit 3
    fi

    if [[ ${#targets[@]} -eq 0 ]]; then
        log_warn "No se encontraron archivos Office para convertir en: $INPUT_PATH"
        exit 0
    fi

    log_info "Archivos a procesar: ${#targets[@]}"
    local ok=0 fail=0 skip=0

    for file in "${targets[@]}"; do
        local out_dir="${OUTPUT_DIR:-}"
        _convert_single_to_odf "$file" "$out_dir"
        local rc=$?
        case $rc in
            0) ((ok++)) ;;
            2) ((skip++)) ;;
            *) ((fail++)) ;;
        esac
    done

    echo ""
    log_info "Resumen → Exitosos: ${ok} | Omitidos: ${skip} | Fallidos: ${fail}"
    _handle_originals_odf
}

# ---------------------------------------------------------------------------
# _collect_odf_targets()
# Busca recursivamente archivos Office en un directorio y llena el array dado.
# Argumentos: $1=nombre_del_array (nameref), $2=directorio_raíz
# ---------------------------------------------------------------------------
_collect_odf_targets() {
    local -n _arr="$1"
    local root_dir="$2"

    # Construir predicados de exclusión para find
    local -a prune_args=()
    for excl in "${EXCLUDE_DIRS[@]}"; do
        prune_args+=(-name "$excl" -o)
    done

    # Construir lista de extensiones a buscar
    local -a name_args=()
    if [[ -n "${FORMATS_FILTER:-}" ]]; then
        IFS=',' read -ra requested <<< "$FORMATS_FILTER"
    else
        # Extraer todas las extensiones de entrada del mapa
        local -a requested=()
        for entry in "${ODF_CONVERSION_MAP[@]}"; do
            requested+=("${entry%%:*}")
        done
    fi

    for ext in "${requested[@]}"; do
        name_args+=(-iname "*.${ext}" -o)
    done
    # Quitar el último -o
    unset 'name_args[-1]'

    if [[ ${#prune_args[@]} -gt 0 ]]; then
        unset 'prune_args[-1]'
        while IFS= read -r -d '' f; do
            _arr+=("$f")
        done < <(find "$root_dir" \
            \( "${prune_args[@]}" \) -prune -o \
            -type f \( "${name_args[@]}" \) -print0 2>/dev/null)
    else
        while IFS= read -r -d '' f; do
            _arr+=("$f")
        done < <(find "$root_dir" \
            -type f \( "${name_args[@]}" \) -print0 2>/dev/null)
    fi
}

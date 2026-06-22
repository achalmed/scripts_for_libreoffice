#!/usr/bin/env bash
# lib/converter_pdf.sh — Conversión directa a PDF (Office/ODF → PDF)
# Por qué existe: la ruta Office → PDF es más corta que Office → ODF → PDF
# y preserva mejor el layout en documentos con formato complejo.
# Bug #4 corregido: LD_LIBRARY_PATH se detecta dinámicamente en vez de hardcodeado.
# Bug #3 corregido: parallel usa --will-cite para evitar bloqueo interactivo.
# Bug #6 corregido: jobs calculados con nproc en vez de hardcodeado a 4.

_PDF_SUCCESS_LOG=""

# ---------------------------------------------------------------------------
# _init_pdf_session()
# ---------------------------------------------------------------------------
_init_pdf_session() {
    _PDF_SUCCESS_LOG="$(mktemp /tmp/docflow_pdf_XXXXXX)"
    log_debug "Registro de sesión PDF: $_PDF_SUCCESS_LOG"
}

_cleanup_pdf_session() {
    [[ -n "$_PDF_SUCCESS_LOG" && -f "$_PDF_SUCCESS_LOG" ]] && rm -f "$_PDF_SUCCESS_LOG"
}

# ---------------------------------------------------------------------------
# _setup_libreoffice_env()
# Detecta la ruta de LibreOffice dinámicamente para evitar el bug del
# LD_LIBRARY_PATH hardcodeado que fallaba en Arch Linux.
# ---------------------------------------------------------------------------
_setup_libreoffice_env() {
    local lo_program_dir
    lo_program_dir="$(find /usr/lib/libreoffice /opt/libreoffice* \
        -maxdepth 2 -name "soffice.bin" -exec dirname {} \; 2>/dev/null | head -1)"

    if [[ -n "$lo_program_dir" ]]; then
        export LD_LIBRARY_PATH="${lo_program_dir}:${LD_LIBRARY_PATH:-}"
        log_debug "LibreOffice program dir: $lo_program_dir"
    fi
    # En Arch Linux la instalación de paquetes ya configura el LD correcto;
    # si no se encuentra nada, LibreOffice igual funciona sin este export.
}

# ---------------------------------------------------------------------------
# _convert_single_to_pdf()
# Convierte un único archivo a PDF.
# Usa pandoc para .docx/.odt y soffice para el resto.
# Argumentos: $1=archivo_entrada, $2=directorio_salida (opcional)
# Retorna: 0 éxito, 1 error, 2 omitido
# ---------------------------------------------------------------------------
_convert_single_to_pdf() {
    local input_file="$1"
    local out_dir="${2:-$(dirname "$input_file")}"
    local ext="${input_file##*.}"
    ext="${ext,,}"
    local base_name
    base_name="$(basename "${input_file%.*}")"
    local out_file="${out_dir}/${base_name}.pdf"

    if [[ -f "$out_file" && "${OVERWRITE:-false}" != "true" ]]; then
        log_detail "$(basename "$input_file") → ya existe, omitido" "⊘"
        record_result "$input_file" "$out_file" "skip" "to-pdf"
        echo "$input_file" >> "$_PDF_SUCCESS_LOG"
        return 2
    fi

    log_detail "$(basename "$input_file") → $(basename "$out_file")"

    if [[ "${DRY_RUN:-false}" == "true" ]]; then
        log_debug "[dry-run] convertir $input_file a PDF en $out_dir"
        record_result "$input_file" "$out_file" "dry-run" "to-pdf"
        return 0
    fi

    validate_writable_dir "$out_dir" || return 1

    local success=false

    # Pandoc maneja mejor .docx y .odt que LibreOffice para texto
    if _ext_in_list "$ext" "${PANDOC_TO_PDF_EXTS[@]}" && command -v pandoc &>/dev/null; then
        log_debug "Usando pandoc para: $(basename "$input_file")"
        if pandoc "$input_file" -o "$out_file" ${PANDOC_EXTRA_ARGS:-} 2>/dev/null; then
            success=true
        fi
    fi

    # Fallback a LibreOffice (o ruta principal para presentaciones/hojas)
    if [[ "$success" == "false" ]]; then
        log_debug "Usando soffice para: $(basename "$input_file")"
        if soffice --headless --convert-to pdf --outdir "$out_dir" "$input_file" &>/dev/null; then
            success=true
        fi
    fi

    if [[ "$success" == "true" ]]; then
        log_detail "✓ $(basename "$out_file")" "✓"
        record_result "$input_file" "$out_file" "ok" "to-pdf"
        echo "$input_file" >> "$_PDF_SUCCESS_LOG"
        return 0
    else
        log_error "Falló: $(basename "$input_file")"
        record_result "$input_file" "$out_file" "fail" "to-pdf"
        return 1
    fi
}

# ---------------------------------------------------------------------------
# _ext_in_list()
# Comprueba si una extensión está en un array de extensiones.
# Argumentos: $1=extensión, $2...=lista de extensiones
# ---------------------------------------------------------------------------
_ext_in_list() {
    local target="${1,,}"
    shift
    for item in "$@"; do
        [[ "${item,,}" == "$target" ]] && return 0
    done
    return 1
}

# ---------------------------------------------------------------------------
# _collect_pdf_targets()
# Busca recursivamente archivos convertibles a PDF.
# Argumentos: $1=nombre_array (nameref), $2=directorio_raíz
# ---------------------------------------------------------------------------
_collect_pdf_targets() {
    local -n _parr="$1"
    local root_dir="$2"

    local -a all_exts=(
        "${PANDOC_TO_PDF_EXTS[@]}"
        "${SOFFICE_TO_PDF_EXTS[@]}"
    )

    local -a name_args=()
    if [[ -n "${FORMATS_FILTER:-}" ]]; then
        IFS=',' read -ra requested <<< "$FORMATS_FILTER"
    else
        requested=("${all_exts[@]}")
    fi

    for ext in "${requested[@]}"; do
        name_args+=(-iname "*.${ext}" -o)
    done
    unset 'name_args[-1]'

    while IFS= read -r -d '' f; do
        _parr+=("$f")
    done < <(find "$root_dir" -type f \( "${name_args[@]}" \) -print0 2>/dev/null)
}

# ---------------------------------------------------------------------------
# _run_parallel_pdf()
# Convierte en paralelo usando xargs -P (más portable que GNU parallel,
# sin el bloqueo interactivo de --will-cite).
# Argumentos: $1=archivo_con_lista_de_rutas
# ---------------------------------------------------------------------------
_run_parallel_pdf() {
    local file_list="$1"
    local out_dir="${OUTPUT_DIR:-}"
    local jobs="${PARALLEL_JOBS:-$(nproc)}"

    log_info "Convirtiendo en paralelo con ${jobs} workers..."

    # Exportar variables y funciones necesarias en los subshells de xargs
    export -f _convert_single_to_pdf _ext_in_list log_detail log_debug \
               log_error record_result validate_writable_dir _setup_libreoffice_env
    export OVERWRITE DRY_RUN PANDOC_EXTRA_ARGS VERBOSE
    export _PDF_SUCCESS_LOG
    export -a PANDOC_TO_PDF_EXTS SOFFICE_TO_PDF_EXTS
    # Exportar colores
    export C_BLUE C_RESET C_CYAN C_RED C_YELLOW C_GREEN C_BOLD
    # Exportar LOG_FILE para que record_result pueda escribir
    export LOG_FILE LOG_RESULTS

    _setup_libreoffice_env

    xargs -0 -P "$jobs" -I{} bash -c \
        '_convert_single_to_pdf "$@"' _ {} "$out_dir" \
        < "$file_list"
}

# ---------------------------------------------------------------------------
# _handle_originals_pdf()
# Gestiona originales después de conversión PDF.
# ---------------------------------------------------------------------------
_handle_originals_pdf() {
    [[ ! -f "$_PDF_SUCCESS_LOG" ]] && return 0
    local count
    count="$(wc -l < "$_PDF_SUCCESS_LOG")"
    [[ "$count" -eq 0 ]] && return 0

    if [[ "${BACKUP_ORIGINALS:-false}" == "true" ]]; then
        log_header "MOVIENDO ORIGINALES A _backup/"
        while IFS= read -r file; do
            local bdir
            bdir="$(dirname "$file")/${BACKUP_SUFFIX:-_backup}"
            mkdir -p "$bdir"
            [[ "${DRY_RUN:-false}" == "true" ]] \
                && log_detail "[dry-run] mv $(basename "$file")" "↦" \
                || { mv "$file" "$bdir/" && log_detail "↦ $(basename "$file")" "↦"; }
        done < "$_PDF_SUCCESS_LOG"
        return 0
    fi

    if [[ "${DELETE_ORIGINALS:-false}" == "true" ]]; then
        log_header "ELIMINACIÓN DE ORIGINALES"
        if confirm_action "¿Eliminar ${count} archivo(s) original(es)?"; then
            while IFS= read -r file; do
                [[ "${DRY_RUN:-false}" == "true" ]] \
                    && log_detail "[dry-run] rm $(basename "$file")" "🗑" \
                    || { rm "$file" && log_detail "🗑 Eliminado: $(basename "$file")" "🗑"; }
            done < "$_PDF_SUCCESS_LOG"
        else
            log_warn "Originales conservados."
        fi
    fi
}

# ---------------------------------------------------------------------------
# run_to_pdf()
# Punto de entrada público del comando to-pdf.
# ---------------------------------------------------------------------------
run_to_pdf() {
    log_header "CONVERSIÓN: Office/ODF → PDF"
    validate_dependencies soffice || exit 5
    _setup_libreoffice_env
    _init_pdf_session
    trap _cleanup_pdf_session EXIT

    local -a targets=()

    if [[ -f "${INPUT_PATH:-}" ]]; then
        targets+=("$INPUT_PATH")
    elif [[ -d "${INPUT_PATH:-}" ]]; then
        _collect_pdf_targets targets "$INPUT_PATH"
    else
        log_error "Ruta inválida o no especificada: '${INPUT_PATH:-}'"
        echo "  Usa: docflow to-pdf -i <archivo_o_directorio>"
        exit 3
    fi

    if [[ ${#targets[@]} -eq 0 ]]; then
        log_warn "No se encontraron archivos convertibles en: $INPUT_PATH"
        exit 0
    fi

    log_info "Archivos encontrados: ${#targets[@]}"

    local ok=0 fail=0 skip=0

    # Para más de 3 archivos, usar procesamiento paralelo
    if [[ ${#targets[@]} -gt 3 ]]; then
        local tmp_list
        tmp_list="$(mktemp /tmp/docflow_list_XXXXXX)"
        printf '%s\0' "${targets[@]}" > "$tmp_list"
        _run_parallel_pdf "$tmp_list"
        rm -f "$tmp_list"
        # El conteo exacto viene del log
        ok="$(grep -c "^" "$_PDF_SUCCESS_LOG" 2>/dev/null || echo 0)"
        fail=$(( ${#targets[@]} - ok ))
    else
        for file in "${targets[@]}"; do
            _convert_single_to_pdf "$file" "${OUTPUT_DIR:-}"
            local rc=$?
            case $rc in
                0) ((ok++)) ;;
                2) ((skip++)) ;;
                *) ((fail++)) ;;
            esac
        done
    fi

    echo ""
    log_info "Resumen → Exitosos: ${ok} | Omitidos: ${skip} | Fallidos: ${fail}"
    _handle_originals_pdf
}

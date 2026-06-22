#!/usr/bin/env bash
# lib/converter_md.sh — Conversión Office/ODF → Markdown con imágenes
# Por qué existe: este módulo refactoriza el script doc2md.sh original,
# corrigiendo la corrección de rutas de imágenes y añadiendo soporte
# para integrar el resultado con el flujo unificado de docflow.

_MD_SUCCESS_LOG=""

_init_md_session() {
    _MD_SUCCESS_LOG="$(mktemp /tmp/docflow_md_XXXXXX)"
}

_cleanup_md_session() {
    [[ -n "$_MD_SUCCESS_LOG" && -f "$_MD_SUCCESS_LOG" ]] && rm -f "$_MD_SUCCESS_LOG"
}

# ---------------------------------------------------------------------------
# _fix_image_paths()
# Corrige las rutas de imágenes en el .md generado por pandoc.
# Pandoc escribe rutas absolutas; necesitamos relativas al .md.
# Argumentos: $1=archivo .md, $2=directorio de imágenes relativo al .md
# ---------------------------------------------------------------------------
_fix_image_paths() {
    local md_file="$1"
    local img_dir_name="$2"

    # Python3 para la sustitución porque sed tiene diferencias entre GNU/BSD
    # que hacen frágil el manejo de rutas con caracteres especiales.
    python3 - "$md_file" "$img_dir_name" <<'PYEOF'
import sys, re, pathlib

md_path   = pathlib.Path(sys.argv[1])
img_dir   = sys.argv[2]
content   = md_path.read_text(encoding="utf-8", errors="replace")

# Reemplaza cualquier ruta absoluta de imagen por una ruta relativa
# que apunte a <nombre>_files/figure-md/
content = re.sub(
    r'!\[([^\]]*)\]\(([^)]+)\)',
    lambda m: f'![{m.group(1)}]({img_dir}/{pathlib.Path(m.group(2)).name})',
    content
)

md_path.write_text(content, encoding="utf-8")
PYEOF
}

# ---------------------------------------------------------------------------
# _optimize_images()
# Convierte imágenes extraídas al formato/calidad solicitados (si imagemagick está).
# Argumentos: $1=directorio de imágenes
# ---------------------------------------------------------------------------
_optimize_images() {
    local img_dir="$1"
    [[ ! -d "$img_dir" ]] && return 0
    command -v convert &>/dev/null || { log_debug "imagemagick no instalado, imágenes sin optimizar"; return 0; }

    local target_fmt="${IMG_FORMAT:-png}"
    local quality="${IMG_QUALITY:-90}"

    while IFS= read -r -d '' img; do
        local ext="${img##*.}"
        ext="${ext,,}"
        [[ "$ext" == "$target_fmt" ]] && continue

        local new_img="${img%.*}.${target_fmt}"
        if [[ "${DRY_RUN:-false}" != "true" ]]; then
            convert "$img" -quality "$quality" "$new_img" 2>/dev/null && rm -f "$img"
        fi
        log_debug "Imagen convertida: $(basename "$img") → $(basename "$new_img")"
    done < <(find "$img_dir" -type f \( -iname "*.png" -o -iname "*.jpg" -o -iname "*.jpeg" -o -iname "*.webp" \) -print0)
}

# ---------------------------------------------------------------------------
# _convert_single_to_md()
# Convierte un único archivo Office/ODF a Markdown.
# Argumentos: $1=archivo_entrada, $2=directorio_salida (opcional)
# Retorna: 0 éxito, 1 error, 2 omitido
# ---------------------------------------------------------------------------
_convert_single_to_md() {
    local input_file="$1"
    local out_dir="${2:-$(dirname "$input_file")}"
    local ext="${input_file##*.}"
    ext="${ext,,}"
    local base_name
    base_name="$(basename "${input_file%.*}")"
    local md_file="${out_dir}/${base_name}.md"
    local img_dir_name="${base_name}_files/figure-md"
    local img_dir="${out_dir}/${img_dir_name}"

    if [[ -f "$md_file" && "${OVERWRITE:-false}" != "true" ]]; then
        log_detail "$(basename "$input_file") → ya existe, omitido" "⊘"
        record_result "$input_file" "$md_file" "skip" "to-md"
        echo "$input_file" >> "$_MD_SUCCESS_LOG"
        return 2
    fi

    log_detail "$(basename "$input_file") → $(basename "$md_file")"

    if [[ "${DRY_RUN:-false}" == "true" ]]; then
        record_result "$input_file" "$md_file" "dry-run" "to-md"
        return 0
    fi

    validate_writable_dir "$out_dir" || return 1

    # Para .pptx y .odp, LibreOffice convierte primero a .odp o .pptx temporal
    # porque pandoc no maneja bien todas las versiones de presentaciones.
    local actual_input="$input_file"
    local tmp_converted=""

    if [[ "$ext" == "pptx" ]] && command -v soffice &>/dev/null; then
        local tmp_dir
        tmp_dir="$(mktemp -d /tmp/docflow_pptx_XXXXXX)"
        if soffice --headless --convert-to odp --outdir "$tmp_dir" "$input_file" &>/dev/null; then
            tmp_converted="${tmp_dir}/${base_name}.odp"
            actual_input="$tmp_converted"
            log_debug "pptx→odp temporal: $tmp_converted"
        fi
    fi

    # Ejecutar pandoc
    local pandoc_cmd=(
        pandoc
        "$actual_input"
        -o "$md_file"
        --extract-media="$img_dir"
        --wrap=none
    )
    [[ -n "${PANDOC_EXTRA_ARGS:-}" ]] && pandoc_cmd+=($PANDOC_EXTRA_ARGS)

    if "${pandoc_cmd[@]}" 2>/dev/null; then
        _fix_image_paths "$md_file" "$img_dir_name"
        _optimize_images "$img_dir"
        log_detail "✓ $(basename "$md_file")" "✓"
        record_result "$input_file" "$md_file" "ok" "to-md"
        echo "$input_file" >> "$_MD_SUCCESS_LOG"
        [[ -n "$tmp_converted" ]] && rm -rf "$(dirname "$tmp_converted")"
        return 0
    else
        log_error "Falló pandoc: $(basename "$input_file")"
        record_result "$input_file" "$md_file" "fail" "to-md"
        [[ -n "$tmp_converted" ]] && rm -rf "$(dirname "$tmp_converted")"
        return 1
    fi
}

# ---------------------------------------------------------------------------
# _collect_md_targets()
# Busca archivos convertibles a Markdown, aplicando exclusiones.
# Argumentos: $1=nombre_array (nameref), $2=directorio_raíz
# ---------------------------------------------------------------------------
_collect_md_targets() {
    local -n _marr="$1"
    local root_dir="$2"

    local supported_exts=("docx" "odt" "pptx" "odp")
    local -a requested_exts=()

    if [[ -n "${FORMATS_FILTER:-}" ]]; then
        IFS=',' read -ra requested_exts <<< "$FORMATS_FILTER"
    else
        requested_exts=("${supported_exts[@]}")
    fi

    # Construir predicados -name para find
    local -a name_args=()
    for ext in "${requested_exts[@]}"; do
        name_args+=(-iname "*.${ext}" -o)
    done
    unset 'name_args[-1]'

    # Construir predicados de poda (prune) para directorios excluidos
    local -a prune_args=()
    if [[ ${#EXCLUDE_DIRS[@]} -gt 0 ]]; then
        prune_args+=(\()
        for excl in "${EXCLUDE_DIRS[@]}"; do
            if [[ "$excl" == /* ]]; then
                prune_args+=(-path "$excl" -o)
            else
                prune_args+=(-name "$excl" -o)
            fi
        done
        unset 'prune_args[-1]'
        prune_args+=(\))
    fi

    if [[ ${#prune_args[@]} -gt 0 ]]; then
        while IFS= read -r -d '' f; do
            _marr+=("$f")
        done < <(find "$root_dir" \
            "${prune_args[@]}" -prune -o \
            -type f \( "${name_args[@]}" \) -print0 2>/dev/null)
    else
        while IFS= read -r -d '' f; do
            _marr+=("$f")
        done < <(find "$root_dir" \
            -type f \( "${name_args[@]}" \) -print0 2>/dev/null)
    fi
}

# ---------------------------------------------------------------------------
# _handle_originals_md()
# ---------------------------------------------------------------------------
_handle_originals_md() {
    [[ ! -f "$_MD_SUCCESS_LOG" ]] && return 0
    local count
    count="$(wc -l < "$_MD_SUCCESS_LOG")"
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
        done < "$_MD_SUCCESS_LOG"
        return 0
    fi

    if [[ "${DELETE_ORIGINALS:-false}" == "true" ]]; then
        log_header "ELIMINACIÓN DE ORIGINALES"
        if confirm_action "¿Eliminar ${count} archivo(s) original(es)?"; then
            while IFS= read -r file; do
                [[ "${DRY_RUN:-false}" == "true" ]] \
                    && log_detail "[dry-run] rm $(basename "$file")" "🗑" \
                    || { rm "$file" && log_detail "🗑 Eliminado: $(basename "$file")" "🗑"; }
            done < "$_MD_SUCCESS_LOG"
        else
            log_warn "Originales conservados."
        fi
    fi
}

# ---------------------------------------------------------------------------
# run_to_md()
# Punto de entrada público del comando to-md.
# ---------------------------------------------------------------------------
run_to_md() {
    log_header "CONVERSIÓN: Office/ODF → Markdown"
    validate_dependencies pandoc python3 || exit 5

    _init_md_session
    trap _cleanup_md_session EXIT

    local -a targets=()

    if [[ -f "${INPUT_PATH:-}" ]]; then
        targets+=("$INPUT_PATH")
    elif [[ -d "${INPUT_PATH:-}" ]]; then
        _collect_md_targets targets "$INPUT_PATH"
    else
        log_error "Ruta inválida o no especificada: '${INPUT_PATH:-}'"
        echo "  Usa: docflow to-md -i <archivo_o_directorio>"
        exit 3
    fi

    if [[ ${#targets[@]} -eq 0 ]]; then
        log_warn "No se encontraron archivos para convertir a Markdown en: $INPUT_PATH"
        exit 0
    fi

    log_info "Archivos a convertir: ${#targets[@]}"
    local ok=0 fail=0 skip=0

    for file in "${targets[@]}"; do
        _convert_single_to_md "$file" "${OUTPUT_DIR:-}"
        local rc=$?
        case $rc in
            0) ((ok++)) ;;
            2) ((skip++)) ;;
            *) ((fail++)) ;;
        esac
    done

    echo ""
    log_info "Resumen → Exitosos: ${ok} | Omitidos: ${skip} | Fallidos: ${fail}"
    _handle_originals_md
}

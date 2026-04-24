#!/usr/bin/env bash
# =============================================================================
#  doc2md.sh — Conversor recursivo de documentos Office a Markdown
#  Soporta: .docx  .odt  .pptx  .odp
#
#  Estructura de salida por cada archivo convertido:
#    <destino>/<ruta_relativa>/index.md
#    <destino>/<ruta_relativa>/index_files/figure-md/fig-001.png ...
#
#  Autor  : Edison Achalma <achalmaedison>
#  Repo   : https://github.com/achalmed/doc2md
#  Versión: 1.0.0
#  Licencia: MIT
# =============================================================================

set -euo pipefail

# ---------------------------------------------------------------------------
# COLORES para la terminal
# ---------------------------------------------------------------------------
RED='\033[0;31m'; GREEN='\033[0;32m'; YELLOW='\033[1;33m'
BLUE='\033[0;34m'; CYAN='\033[0;36m'; BOLD='\033[1m'; RESET='\033[0m'

# ---------------------------------------------------------------------------
# VALORES POR DEFECTO
# ---------------------------------------------------------------------------
INPUT_DIR=""
OUTPUT_DIR=""
VERBOSE=false
DRY_RUN=false
OVERWRITE=false
FORMATS=("docx" "odt" "pptx" "odp")
LOG_FILE=""
PANDOC_EXTRA_ARGS=""
IMG_FORMAT="png"          # png | jpg | webp
IMG_QUALITY=90            # 1-100 (para jpg/webp)
FLATTEN=false             # Si true, no replica estructura de directorios
SUMMARY_FILE=""           # Archivo resumen al final

# Contadores globales
COUNT_OK=0
COUNT_SKIP=0
COUNT_ERR=0
COUNT_TOTAL=0
declare -a ERRORS=()

# ---------------------------------------------------------------------------
# FUNCIONES DE UTILIDAD
# ---------------------------------------------------------------------------
log()      { echo -e "${BLUE}[INFO]${RESET}  $*"; }
log_ok()   { echo -e "${GREEN}[OK]${RESET}    $*"; }
log_warn() { echo -e "${YELLOW}[WARN]${RESET}  $*"; }
log_err()  { echo -e "${RED}[ERROR]${RESET} $*" >&2; }
log_dbg()  { $VERBOSE && echo -e "${CYAN}[DEBUG]${RESET} $*" || true; }
log_sep()  { echo -e "${BOLD}──────────────────────────────────────────────────${RESET}"; }

die() { log_err "$*"; exit 1; }

usage() {
cat << EOF
${BOLD}USO${RESET}
    $(basename "$0") [OPCIONES] -i <DIRECTORIO_ENTRADA> -o <DIRECTORIO_SALIDA>

${BOLD}OPCIONES REQUERIDAS${RESET}
    -i, --input   <dir>     Directorio fuente (se escanea recursivamente)
    -o, --output  <dir>     Directorio de destino para los .md generados

${BOLD}OPCIONES GENERALES${RESET}
    -f, --formats <lista>   Formatos a convertir, separados por coma
                            Por defecto: docx,odt,pptx,odp
                            Ejemplo: --formats docx,odt

    -w, --overwrite         Sobreescribir archivos .md ya existentes
                            (por defecto se omiten)

    --flatten               No replicar estructura de subdirectorios;
                            todos los .md van al raíz del directorio de salida

    --img-format <fmt>      Formato de las imágenes extraídas: png | jpg | webp
                            Por defecto: png

    --img-quality <n>       Calidad de imagen para jpg/webp (1-100)
                            Por defecto: 90

    --pandoc-args <args>    Argumentos extra para pandoc (entre comillas)
                            Ejemplo: '--wrap=none --toc'

    -l, --log <archivo>     Guardar log en un archivo (además de stdout)
    -s, --summary <archivo> Generar archivo resumen de la conversión

    -v, --verbose           Salida detallada
    -n, --dry-run           Simular sin escribir nada
    -h, --help              Mostrar esta ayuda

${BOLD}EJEMPLOS${RESET}
    # Conversión básica
    $(basename "$0") -i ~/Documentos -o ~/Markdown

    # Solo .docx y .odt, sobreescribir existentes
    $(basename "$0") -i ~/Documentos -o ~/Markdown -f docx,odt -w

    # Con log y resumen
    $(basename "$0") -i ./docs -o ./md -l conversion.log -s resumen.txt -v

    # Imágenes en JPEG de alta calidad
    $(basename "$0") -i ./docs -o ./md --img-format jpg --img-quality 95

${BOLD}ESTRUCTURA DE SALIDA${RESET}
    Para cada archivo convertido se genera:
        <salida>/<ruta_relativa>/
            index.md                     ← contenido Markdown
            index_files/
                figure-md/
                    fig-001.png          ← imágenes extraídas y renombradas
                    fig-002.png
                    ...

${BOLD}DEPENDENCIAS${RESET}
    Requeridas : pandoc, python3
    Opcionales : libreoffice (fallback para pptx/odp), imagemagick (optimización)
EOF
    exit 0
}

# ---------------------------------------------------------------------------
# VERIFICAR DEPENDENCIAS
# ---------------------------------------------------------------------------
check_deps() {
    local missing=()
    for cmd in pandoc python3; do
        command -v "$cmd" &>/dev/null || missing+=("$cmd")
    done
    if [[ ${#missing[@]} -gt 0 ]]; then
        die "Dependencias faltantes: ${missing[*]}\n  Instala con: sudo pacman -S ${missing[*]}"
    fi
    # Avisos opcionales
    command -v libreoffice &>/dev/null || \
        log_warn "libreoffice no encontrado — fallback para pptx/odp puede fallar"
    command -v magick &>/dev/null || command -v convert &>/dev/null || \
        log_warn "imagemagick no encontrado — optimización de imágenes desactivada"
    log_dbg "Pandoc: $(pandoc --version | head -1)"
}

# ---------------------------------------------------------------------------
# PARSEAR ARGUMENTOS
# ---------------------------------------------------------------------------
parse_args() {
    [[ $# -eq 0 ]] && usage

    while [[ $# -gt 0 ]]; do
        case "$1" in
            -i|--input)       INPUT_DIR="$2";          shift 2 ;;
            -o|--output)      OUTPUT_DIR="$2";         shift 2 ;;
            -f|--formats)
                IFS=',' read -ra FORMATS <<< "$2";     shift 2 ;;
            -w|--overwrite)   OVERWRITE=true;          shift   ;;
            --flatten)        FLATTEN=true;            shift   ;;
            --img-format)     IMG_FORMAT="$2";         shift 2 ;;
            --img-quality)    IMG_QUALITY="$2";        shift 2 ;;
            --pandoc-args)    PANDOC_EXTRA_ARGS="$2";  shift 2 ;;
            -l|--log)         LOG_FILE="$2";           shift 2 ;;
            -s|--summary)     SUMMARY_FILE="$2";       shift 2 ;;
            -v|--verbose)     VERBOSE=true;            shift   ;;
            -n|--dry-run)     DRY_RUN=true;            shift   ;;
            -h|--help)        usage ;;
            *) die "Opción desconocida: $1 (usa -h para ayuda)" ;;
        esac
    done

    # Validaciones
    [[ -z "$INPUT_DIR"  ]] && die "Falta -i/--input"
    [[ -z "$OUTPUT_DIR" ]] && die "Falta -o/--output"
    [[ ! -d "$INPUT_DIR" ]] && die "Directorio de entrada no existe: $INPUT_DIR"
    [[ "$IMG_FORMAT" =~ ^(png|jpg|webp)$ ]] || \
        die "--img-format debe ser png, jpg o webp"
    [[ "$IMG_QUALITY" =~ ^[0-9]+$ ]] && \
        [[ "$IMG_QUALITY" -ge 1 && "$IMG_QUALITY" -le 100 ]] || \
        die "--img-quality debe ser un número entre 1 y 100"

    # Redirigir también a log si se pidió
    if [[ -n "$LOG_FILE" ]]; then
        exec > >(tee -a "$LOG_FILE") 2>&1
    fi
}

# ---------------------------------------------------------------------------
# OBTENER RUTA DE SALIDA para un archivo fuente dado
# ---------------------------------------------------------------------------
get_output_dir() {
    local src_file="$1"          # ruta absoluta del fuente
    local src_base               # nombre sin extensión
    src_base=$(basename "$src_file")
    src_base="${src_base%.*}"

    if $FLATTEN; then
        echo "${OUTPUT_DIR}/${src_base}"
    else
        # Replicar estructura relativa al INPUT_DIR
        local rel_dir
        rel_dir=$(dirname "${src_file#$INPUT_DIR/}")
        if [[ "$rel_dir" == "." ]]; then
            echo "${OUTPUT_DIR}/${src_base}"
        else
            echo "${OUTPUT_DIR}/${rel_dir}/${src_base}"
        fi
    fi
}

# ---------------------------------------------------------------------------
# RENOMBRAR IMÁGENES al formato fig-NNN.ext y dejarlas en figure-md/
# ---------------------------------------------------------------------------
rename_figures() {
    local figures_dir="$1"
    local counter=0

    # Recoger todas las imágenes generadas por pandoc
    local tmplist
    tmplist=$(mktemp)
    find "$figures_dir" -maxdepth 1 -type f \
        \( -iname "*.png" -o -iname "*.jpg" -o -iname "*.jpeg" \
           -o -iname "*.gif" -o -iname "*.svg" -o -iname "*.webp" \) \
        | sort > "$tmplist"

    while IFS= read -r img_path; do
        counter=$(( counter + 1 ))
        local ext="${img_path##*.}"
        local new_name
        printf -v new_name "fig-%03d.%s" "$counter" "${ext,,}"
        local new_path="${figures_dir}/${new_name}"

        if [[ "$img_path" != "$new_path" ]]; then
            mv "$img_path" "$new_path"
            log_dbg "  Renombrado: $(basename "$img_path") → ${new_name}"
        fi
    done < "$tmplist"
    rm -f "$tmplist"

    echo "$counter"   # devuelve número de figuras procesadas
}

# ---------------------------------------------------------------------------
# OPTIMIZAR IMAGEN con imagemagick (si disponible y no es PNG puro)
# ---------------------------------------------------------------------------
optimize_image() {
    local img="$1"
    local magick_bin=""
    command -v magick  &>/dev/null && magick_bin="magick"
    command -v convert &>/dev/null && magick_bin="convert"
    [[ -z "$magick_bin" ]] && return 0

    case "$IMG_FORMAT" in
        jpg)
            $magick_bin "$img" -quality "$IMG_QUALITY" \
                "${img%.*}.jpg" 2>/dev/null && \
            [[ "${img%.*}.jpg" != "$img" ]] && rm -f "$img" || true
            ;;
        webp)
            $magick_bin "$img" -quality "$IMG_QUALITY" \
                "${img%.*}.webp" 2>/dev/null && \
            [[ "${img%.*}.webp" != "$img" ]] && rm -f "$img" || true
            ;;
        *) : ;;   # PNG no necesita conversión extra
    esac
}

# ---------------------------------------------------------------------------
# CORREGIR rutas de imagen en el .md generado
# Los paths que pandoc escribe pueden ser absolutos o relativos al tmpdir.
# Los normalizamos a: index_files/figure-md/fig-NNN.ext
# ---------------------------------------------------------------------------
fix_image_paths() {
    local md_file="$1"
    local figures_rel="index_files/figure-md"

    python3 - "$md_file" "$figures_rel" << 'PYEOF'
import sys, re, os

md_file     = sys.argv[1]
figures_rel = sys.argv[2]

with open(md_file, 'r', encoding='utf-8') as f:
    content = f.read()

# Patrones pandoc: ![alt](path) y ![alt](path "title")
img_pattern = re.compile(
    r'!\[([^\]]*)\]\(([^)]+?)(\s+"[^"]*")?\)',
    re.MULTILINE
)

def replace_img(m):
    alt   = m.group(1)
    path  = m.group(2).strip()
    title = m.group(3) or ''
    basename = os.path.basename(path)
    new_path = f"{figures_rel}/{basename}"
    return f'![{alt}]({new_path}{title})'

new_content = img_pattern.sub(replace_img, content)

with open(md_file, 'w', encoding='utf-8') as f:
    f.write(new_content)
PYEOF
}

# ---------------------------------------------------------------------------
# CONVERTIR UN SOLO ARCHIVO
# ---------------------------------------------------------------------------
convert_file() {
    local src_file="$1"
    local src_ext="${src_file##*.}"
    src_ext="${src_ext,,}"

    COUNT_TOTAL=$(( COUNT_TOTAL + 1 ))

    # Directorio y archivo de salida
    local out_dir
    out_dir=$(get_output_dir "$src_file")
    local out_md="${out_dir}/index.md"
    local figures_dir="${out_dir}/index_files/figure-md"

    log_sep
    log "Procesando: ${BOLD}$(basename "$src_file")${RESET}"
    log_dbg "  Fuente : $src_file"
    log_dbg "  Destino: $out_md"

    # ── Comprobar si ya existe ──────────────────────────────────────────────
    if [[ -f "$out_md" ]] && ! $OVERWRITE; then
        log_warn "Ya existe (usa -w para sobreescribir): $out_md"
        COUNT_SKIP=$(( COUNT_SKIP + 1 ))
        return 0
    fi

    $DRY_RUN && { log "[DRY-RUN] Se crearía: $out_md"; return 0; }

    # ── Crear directorios ───────────────────────────────────────────────────
    mkdir -p "$figures_dir"

    # ── Construir comando pandoc ────────────────────────────────────────────
    # --extract-media vuelca las imágenes en la carpeta que indiquemos.
    # Usamos un directorio temporal para que pandoc no cree subcarpetas propias
    # y luego movemos todo a figures_dir.
    local tmp_media
    tmp_media=$(mktemp -d)

    local pandoc_opts=(
        --from="$src_ext"
        --to=markdown
        --standalone
        --wrap=none
        --extract-media="$tmp_media"
        --output="$out_md"
    )

    # Añadir argumentos extra del usuario (si existen)
    if [[ -n "$PANDOC_EXTRA_ARGS" ]]; then
        # shellcheck disable=SC2206
        pandoc_opts+=( $PANDOC_EXTRA_ARGS )
    fi

    # ── Ejecutar pandoc ─────────────────────────────────────────────────────
    local pandoc_ok=true
    if ! pandoc "${pandoc_opts[@]}" "$src_file" 2>/tmp/pandoc_stderr.txt; then
        pandoc_ok=false
    fi

    # Para pptx/odp: si pandoc falla, intentar con LibreOffice como pre-conversión
    if ! $pandoc_ok && [[ "$src_ext" =~ ^(pptx|odp)$ ]]; then
        log_warn "Pandoc falló para $src_ext, intentando pre-conversión con LibreOffice..."
        local tmp_odt
        tmp_odt=$(mktemp --suffix=".odt")
        if libreoffice --headless --convert-to odt --outdir "$(dirname "$tmp_odt")" \
               "$src_file" &>/dev/null; then
            local lo_out
            lo_out="$(dirname "$tmp_odt")/$(basename "${src_file%.*}.odt")"
            pandoc_opts[0]="--from=odt"
            if pandoc "${pandoc_opts[@]}" "$lo_out" 2>/tmp/pandoc_stderr.txt; then
                pandoc_ok=true
            fi
            rm -f "$lo_out"
        fi
        rm -f "$tmp_odt"
    fi

    if ! $pandoc_ok; then
        log_err "Falló la conversión de: $src_file"
        log_err "  $(cat /tmp/pandoc_stderr.txt)"
        ERRORS+=("$src_file")
        COUNT_ERR=$(( COUNT_ERR + 1 ))
        rm -rf "$tmp_media"
        return 1
    fi

    # ── Mover imágenes extraídas → figures_dir ──────────────────────────────
    if [[ -d "$tmp_media" ]]; then
        find "$tmp_media" -type f \
            \( -iname "*.png" -o -iname "*.jpg" -o -iname "*.jpeg" \
               -o -iname "*.gif" -o -iname "*.svg" -o -iname "*.webp" \
               -o -iname "*.bmp" -o -iname "*.tiff" \) \
            -exec mv -n {} "$figures_dir/" \; 2>/dev/null || true
        rm -rf "$tmp_media"
    fi

    # ── Renombrar figuras a fig-NNN.ext ─────────────────────────────────────
    local n_figs=0
    if [[ -d "$figures_dir" ]] && \
       [[ -n "$(ls -A "$figures_dir" 2>/dev/null)" ]]; then
        n_figs=$(rename_figures "$figures_dir")

        # Optimización opcional de imágenes
        if [[ "$IMG_FORMAT" != "png" ]]; then
            find "$figures_dir" -type f | while read -r img; do
                optimize_image "$img"
            done
        fi
    fi

    # ── Corregir rutas de imagen en el .md ──────────────────────────────────
    fix_image_paths "$out_md"

    # ── Limpiar figures_dir si quedó vacío ──────────────────────────────────
    if [[ -d "$figures_dir" ]] && \
       [[ -z "$(ls -A "$figures_dir" 2>/dev/null)" ]]; then
        rmdir "$figures_dir" 2>/dev/null || true
        rmdir "${out_dir}/index_files" 2>/dev/null || true
    fi

    COUNT_OK=$(( COUNT_OK + 1 ))
    log_ok "Convertido → $out_md  [figuras: ${n_figs}]"
    return 0
}

# ---------------------------------------------------------------------------
# BUSCAR ARCHIVOS y ejecutar conversiones
# ---------------------------------------------------------------------------
run_conversions() {
    # Construir la expresión -name para find
    local find_args=( "$INPUT_DIR" -type f )
    local name_exprs=()
    for fmt in "${FORMATS[@]}"; do
        name_exprs+=( -iname "*.${fmt}" -o )
    done
    # Quitar el último '-o'
    unset 'name_exprs[${#name_exprs[@]}-1]'

    find_args+=( \( "${name_exprs[@]}" \) )

    log "Escaneando: $INPUT_DIR"
    log "Formatos  : ${FORMATS[*]}"
    log "Destino   : $OUTPUT_DIR"
    $DRY_RUN && log_warn "MODO DRY-RUN — no se escribirá nada"
    log_sep

    while IFS= read -r -d '' src; do
        convert_file "$src" || true
    done < <(find "${find_args[@]}" -print0 | sort -z)
}

# ---------------------------------------------------------------------------
# RESUMEN FINAL
# ---------------------------------------------------------------------------
print_summary() {
    local ts
    ts=$(date '+%Y-%m-%d %H:%M:%S')

    log_sep
    echo -e "${BOLD}RESUMEN DE CONVERSIÓN${RESET}  [$ts]"
    echo -e "  Total encontrados : $COUNT_TOTAL"
    echo -e "  ${GREEN}Convertidos OK    : $COUNT_OK${RESET}"
    echo -e "  ${YELLOW}Omitidos (ya exist): $COUNT_SKIP${RESET}"
    echo -e "  ${RED}Con errores       : $COUNT_ERR${RESET}"

    if [[ ${#ERRORS[@]} -gt 0 ]]; then
        echo -e "\n${RED}Archivos con error:${RESET}"
        for e in "${ERRORS[@]}"; do
            echo "  - $e"
        done
    fi
    log_sep

    # Guardar resumen en archivo si se pidió
    if [[ -n "$SUMMARY_FILE" ]]; then
        {
            echo "=== doc2md — Resumen de conversión ==="
            echo "Fecha     : $ts"
            echo "Entrada   : $INPUT_DIR"
            echo "Salida    : $OUTPUT_DIR"
            echo "Formatos  : ${FORMATS[*]}"
            echo "Total     : $COUNT_TOTAL"
            echo "OK        : $COUNT_OK"
            echo "Omitidos  : $COUNT_SKIP"
            echo "Errores   : $COUNT_ERR"
            if [[ ${#ERRORS[@]} -gt 0 ]]; then
                echo ""
                echo "Archivos con error:"
                for e in "${ERRORS[@]}"; do echo "  $e"; done
            fi
        } > "$SUMMARY_FILE"
        log "Resumen guardado en: $SUMMARY_FILE"
    fi
}

# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------
main() {
    echo -e "${BOLD}${CYAN}"
    echo "  ██████╗  ██████╗  ██████╗██████╗ ███╗   ███╗██████╗ "
    echo "  ██╔══██╗██╔═══██╗██╔════╝╚════██╗████╗ ████║██╔══██╗"
    echo "  ██║  ██║██║   ██║██║      █████╔╝██╔████╔██║██║  ██║"
    echo "  ██║  ██║██║   ██║██║     ██╔═══╝ ██║╚██╔╝██║██║  ██║"
    echo "  ██████╔╝╚██████╔╝╚██████╗███████╗██║ ╚═╝ ██║██████╔╝"
    echo "  ╚═════╝  ╚═════╝  ╚═════╝╚══════╝╚═╝     ╚═╝╚═════╝ "
    echo -e "${RESET}${BOLD}  Conversor recursivo Office → Markdown  v1.0.0${RESET}"
    echo ""

    parse_args "$@"
    check_deps

    # Crear directorio de salida
    $DRY_RUN || mkdir -p "$OUTPUT_DIR"

    run_conversions
    print_summary

    [[ $COUNT_ERR -gt 0 ]] && exit 1 || exit 0
}

main "$@"

#!/usr/bin/env bash
# =============================================================================
#  doc2md.sh — Conversor recursivo de documentos Office a Markdown
#  Soporta: .docx  .odt  .pptx  .odp
#
#  Comportamiento:
#    · El .md se guarda EN EL MISMO DIRECTORIO que el archivo original
#    · El .md recibe el MISMO NOMBRE que el original (solo cambia la extensión)
#    · Las imágenes se extraen en <nombre>_files/figure-md/ junto al .md
#    · La estructura de carpetas NO se modifica ni replica
#
#  Ejemplos de salida:
#    ideas/2017.../index.odt  →  ideas/2017.../index.md
#                                ideas/2017.../index_files/figure-md/fig-001.png
#    docs/informe.docx        →  docs/informe.md
#                                docs/informe_files/figure-md/fig-001.png
#
#  Autor   : Edison Achalma — github.com/achalmed
#  Script  : ~/Documents/scripts_for_libreoffice/script_convert_doc_to_md/doc2md.sh
#  Versión : 2.0.0
#  Licencia: MIT
# =============================================================================

set -euo pipefail

# ---------------------------------------------------------------------------
# DIRECTORIO DONDE VIVE ESTE SCRIPT (siempre el mismo lugar)
# ---------------------------------------------------------------------------
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SCRIPT_NAME="$(basename "${BASH_SOURCE[0]}")"

# ---------------------------------------------------------------------------
# COLORES
# ---------------------------------------------------------------------------
RED='\033[0;31m'; GREEN='\033[0;32m'; YELLOW='\033[1;33m'
BLUE='\033[0;34m'; CYAN='\033[0;36m'; BOLD='\033[1m'; RESET='\033[0m'

# ---------------------------------------------------------------------------
# VALORES POR DEFECTO
# ---------------------------------------------------------------------------
TARGET_DIR=""               # Directorio a escanear (requerido)
VERBOSE=false
DRY_RUN=false
OVERWRITE=false
DELETE_CONVERTED=false      # Eliminar SOLO los archivos Office ya convertidos
FORMATS=("docx" "odt" "pptx" "odp")
LOG_FILE=""
PANDOC_EXTRA_ARGS=""
IMG_FORMAT="png"            # png | jpg | webp
IMG_QUALITY=90              # 1-100 (para jpg/webp)
SUMMARY_FILE=""

# ── Carpetas excluidas por defecto ──────────────────────────────────────────
# Edita esta lista para agregar exclusiones permanentes.
# Puedes usar nombres de carpeta simples o rutas absolutas.
# También se pueden añadir en tiempo de ejecución con --exclude.
declare -a EXCLUDE_DIRS=(
    ".git"
    ".svn"
    "node_modules"
    "__pycache__"
    ".trash"
    "Trash"
    ".Trash"
)

# Contadores globales
COUNT_OK=0
COUNT_SKIP=0
COUNT_ERR=0
COUNT_DELETE=0
COUNT_TOTAL=0
declare -a ERRORS=()
declare -a CONVERTED_FILES=()   # lista de archivos Office convertidos con éxito

# ---------------------------------------------------------------------------
# FUNCIONES DE LOG
# ---------------------------------------------------------------------------
log()      { echo -e "${BLUE}[INFO]${RESET}  $*"; }
log_ok()   { echo -e "${GREEN}[OK]${RESET}    $*"; }
log_warn() { echo -e "${YELLOW}[WARN]${RESET}  $*"; }
log_err()  { echo -e "${RED}[ERROR]${RESET} $*" >&2; }
log_dbg()  { $VERBOSE && echo -e "${CYAN}[DEBUG]${RESET} $*" || true; }
log_sep()  { echo -e "${BOLD}──────────────────────────────────────────────${RESET}"; }

die() { log_err "$*"; exit 1; }

# ---------------------------------------------------------------------------
# AYUDA
# ---------------------------------------------------------------------------
usage() {
cat << EOF
${BOLD}USO${RESET}
    ${SCRIPT_NAME} [OPCIONES] <DIRECTORIO>

    Convierte recursivamente archivos Office a Markdown.
    El .md se genera EN EL MISMO LUGAR que el archivo original,
    con el MISMO NOMBRE (solo cambia la extensión).

${BOLD}ARGUMENTO${RESET}
    <DIRECTORIO>            Directorio raíz a escanear (requerido)

${BOLD}OPCIONES DE CONVERSIÓN${RESET}
    -f, --formats <lista>   Formatos a convertir, separados por coma
                            Por defecto: docx,odt,pptx,odp
                            Ejemplo: -f docx,odt

    -w, --overwrite         Sobreescribir .md ya existentes
                            (por defecto se omiten si ya existen)

    --img-format <fmt>      Formato de las imágenes: png | jpg | webp
                            Por defecto: png

    --img-quality <n>       Calidad de imagen para jpg/webp (1-100)
                            Por defecto: 90

    --pandoc-args <args>    Argumentos extra para pandoc (entre comillas)
                            Ejemplo: --pandoc-args '--wrap=none --toc'

${BOLD}OPCIONES DE EXCLUSIÓN${RESET}
    -e, --exclude <nombre>  Excluir carpeta por nombre o ruta absoluta.
                            Se puede usar múltiples veces:
                              -e carpeta1 -e carpeta2
                            También acepta lista separada por coma:
                              -e "carpeta1,carpeta2"

    --no-default-excludes   No aplicar la lista de exclusiones por defecto
                            (útil si quieres procesar .git u otras carpetas)

${BOLD}OPCIONES DE LIMPIEZA${RESET}
    --delete-converted      Eliminar los archivos Office SOLO si fueron
                            convertidos exitosamente en esta ejecución.
                            NO elimina archivos que ya existían antes,
                            ni archivos que fallaron, ni ningún otro archivo.
                            PRECAUCION: Usa --dry-run primero para verificar.

${BOLD}OPCIONES DE REGISTRO${RESET}
    -l, --log <archivo>     Guardar log en archivo (además de stdout)
    -s, --summary <archivo> Generar archivo resumen de la conversión

${BOLD}OTRAS OPCIONES${RESET}
    -v, --verbose           Salida detallada
    -n, --dry-run           Simular sin escribir nada (muy recomendado antes
                            de usar --delete-converted)
    -h, --help              Mostrar esta ayuda

${BOLD}ESTRUCTURA DE SALIDA${RESET}
    Para cada archivo convertido:

        ideas/2017.../index.odt
        ideas/2017.../index.md                      <- mismo nombre
        ideas/2017.../index_files/figure-md/
            fig-001.png
            fig-002.png

        docs/informe.docx
        docs/informe.md
        docs/informe_files/figure-md/
            fig-001.png

${BOLD}CARPETAS EXCLUIDAS POR DEFECTO${RESET}
    ${EXCLUDE_DIRS[*]}

    Edita la sección EXCLUDE_DIRS en el script para cambiarlas permanentemente.

${BOLD}EJEMPLOS${RESET}
    # Convertir todo en ~/Documents/ideas (salida en el mismo lugar)
    ${SCRIPT_NAME} ~/Documents/ideas

    # Solo .odt y .docx, sobreescribir existentes
    ${SCRIPT_NAME} -f odt,docx -w ~/Documents/ideas

    # Excluir carpetas especificas
    ${SCRIPT_NAME} -e notas -e borradores ~/Documents/ideas

    # Ver que se convertiria sin hacer nada (dry-run)
    ${SCRIPT_NAME} -n -v ~/Documents/ideas

    # Convertir y luego eliminar los originales convertidos
    ${SCRIPT_NAME} --delete-converted ~/Documents/ideas

    # Con log y resumen
    ${SCRIPT_NAME} -l ~/doc2md.log -s ~/resumen.txt ~/Documents/ideas

${BOLD}UBICACION DEL SCRIPT${RESET}
    ${SCRIPT_DIR}/${SCRIPT_NAME}
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
    command -v libreoffice &>/dev/null || \
        log_warn "libreoffice no encontrado — fallback para pptx/odp puede fallar"
    command -v magick  &>/dev/null || \
    command -v convert &>/dev/null || \
        log_dbg "imagemagick no encontrado — optimizacion de imagenes desactivada"
    log_dbg "Pandoc: $(pandoc --version | head -1)"
    log_dbg "Script : ${SCRIPT_DIR}/${SCRIPT_NAME}"
}

# ---------------------------------------------------------------------------
# PARSEAR ARGUMENTOS
# ---------------------------------------------------------------------------
parse_args() {
    [[ $# -eq 0 ]] && usage

    local extra_excludes=()
    local no_default_excludes=false

    while [[ $# -gt 0 ]]; do
        case "$1" in
            -f|--formats)
                IFS=',' read -ra FORMATS <<< "$2"; shift 2 ;;
            -w|--overwrite)
                OVERWRITE=true; shift ;;
            --img-format)
                IMG_FORMAT="$2"; shift 2 ;;
            --img-quality)
                IMG_QUALITY="$2"; shift 2 ;;
            --pandoc-args)
                PANDOC_EXTRA_ARGS="$2"; shift 2 ;;
            -e|--exclude)
                IFS=',' read -ra _ex <<< "$2"
                extra_excludes+=("${_ex[@]}")
                shift 2 ;;
            --no-default-excludes)
                no_default_excludes=true; shift ;;
            --delete-converted)
                DELETE_CONVERTED=true; shift ;;
            -l|--log)
                LOG_FILE="$2"; shift 2 ;;
            -s|--summary)
                SUMMARY_FILE="$2"; shift 2 ;;
            -v|--verbose)
                VERBOSE=true; shift ;;
            -n|--dry-run)
                DRY_RUN=true; shift ;;
            -h|--help)
                usage ;;
            -*)
                die "Opcion desconocida: $1  (usa -h para ayuda)" ;;
            *)
                [[ -n "$TARGET_DIR" ]] && \
                    die "Demasiados argumentos. Solo se acepta un directorio."
                TARGET_DIR="$1"; shift ;;
        esac
    done

    # Validaciones
    [[ -z "$TARGET_DIR" ]] && \
        die "Falta el directorio a escanear.\n  Uso: ${SCRIPT_NAME} [OPCIONES] <DIRECTORIO>"
    [[ ! -d "$TARGET_DIR" ]] && die "El directorio no existe: $TARGET_DIR"
    TARGET_DIR="$(realpath "$TARGET_DIR")"

    [[ "$IMG_FORMAT" =~ ^(png|jpg|webp)$ ]] || \
        die "--img-format debe ser png, jpg o webp"
    [[ "$IMG_QUALITY" =~ ^[0-9]+$ ]] && \
        [[ "$IMG_QUALITY" -ge 1 && "$IMG_QUALITY" -le 100 ]] || \
        die "--img-quality debe ser un numero entre 1 y 100"

    # Construir lista final de exclusiones
    $no_default_excludes && EXCLUDE_DIRS=()
    EXCLUDE_DIRS+=("${extra_excludes[@]}")

    # Redirigir a log si se pidio
    if [[ -n "$LOG_FILE" ]]; then
        exec > >(tee -a "$LOG_FILE") 2>&1
    fi

    # Advertencia de seguridad para --delete-converted
    if $DELETE_CONVERTED && ! $DRY_RUN; then
        log_warn "──────────────────────────────────────────────────────────"
        log_warn "  --delete-converted activado:"
        log_warn "  Se eliminaran los archivos Office convertidos con exito."
        log_warn "  Usa -n (dry-run) primero si tienes dudas."
        log_warn "──────────────────────────────────────────────────────────"
    fi
}

# ---------------------------------------------------------------------------
# RENOMBRAR IMAGENES → fig-NNN.ext dentro de figures_dir
# ---------------------------------------------------------------------------
rename_figures() {
    local figures_dir="$1"
    local counter=0
    local tmplist
    tmplist=$(mktemp)

    find "$figures_dir" -maxdepth 1 -type f \
        \( -iname "*.png" -o -iname "*.jpg" -o -iname "*.jpeg" \
           -o -iname "*.gif" -o -iname "*.svg" -o -iname "*.webp" \
           -o -iname "*.bmp" -o -iname "*.tiff" \) \
        | sort > "$tmplist"

    while IFS= read -r img_path; do
        counter=$(( counter + 1 ))
        local ext="${img_path##*.}"
        local new_name
        printf -v new_name "fig-%03d.%s" "$counter" "${ext,,}"
        local new_path="${figures_dir}/${new_name}"
        [[ "$img_path" != "$new_path" ]] && mv "$img_path" "$new_path"
        log_dbg "  Figura: $(basename "$img_path") → ${new_name}"
    done < "$tmplist"

    rm -f "$tmplist"
    echo "$counter"
}

# ---------------------------------------------------------------------------
# OPTIMIZAR IMAGEN con imagemagick
# ---------------------------------------------------------------------------
optimize_image() {
    local img="$1"
    local magick_bin=""
    command -v magick  &>/dev/null && magick_bin="magick"
    command -v convert &>/dev/null && magick_bin="${magick_bin:-convert}"
    [[ -z "$magick_bin" ]] && return 0

    case "$IMG_FORMAT" in
        jpg)
            $magick_bin "$img" -quality "$IMG_QUALITY" \
                "${img%.*}.jpg" 2>/dev/null && \
            [[ "${img%.*}.jpg" != "$img" ]] && rm -f "$img" || true ;;
        webp)
            $magick_bin "$img" -quality "$IMG_QUALITY" \
                "${img%.*}.webp" 2>/dev/null && \
            [[ "${img%.*}.webp" != "$img" ]] && rm -f "$img" || true ;;
        *) : ;;
    esac
}

# ---------------------------------------------------------------------------
# CORREGIR rutas de imagen en el .md
# ---------------------------------------------------------------------------
fix_image_paths() {
    local md_file="$1"
    local figures_rel="$2"

    python3 - "$md_file" "$figures_rel" << 'PYEOF'
import sys, re, os

md_file     = sys.argv[1]
figures_rel = sys.argv[2]

with open(md_file, 'r', encoding='utf-8') as f:
    content = f.read()

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
# CONVERTIR UN ARCHIVO
# ---------------------------------------------------------------------------
convert_file() {
    local src_file="$1"
    local src_ext="${src_file##*.}"
    src_ext="${src_ext,,}"

    COUNT_TOTAL=$(( COUNT_TOTAL + 1 ))

    # El .md va en el MISMO DIRECTORIO con el MISMO NOMBRE BASE
    local src_dir
    src_dir=$(dirname "$src_file")
    local src_base
    src_base=$(basename "$src_file")
    local name_no_ext="${src_base%.*}"

    local out_md="${src_dir}/${name_no_ext}.md"
    local figures_dir="${src_dir}/${name_no_ext}_files/figure-md"
    local figures_rel="${name_no_ext}_files/figure-md"

    log_sep
    log "Procesando: ${BOLD}${src_base}${RESET}"
    log_dbg "  Dir    : $src_dir"
    log_dbg "  Salida : $out_md"

    # ¿Ya existe el .md?
    if [[ -f "$out_md" ]] && ! $OVERWRITE; then
        log_warn "Ya existe (usa -w para sobreescribir): $(basename "$out_md")"
        COUNT_SKIP=$(( COUNT_SKIP + 1 ))
        return 0
    fi

    # Dry-run
    if $DRY_RUN; then
        log "[DRY-RUN] Se crearia: $out_md"
        $DELETE_CONVERTED && log "[DRY-RUN] Se eliminaria: $src_file"
        COUNT_OK=$(( COUNT_OK + 1 ))
        CONVERTED_FILES+=("$src_file")
        return 0
    fi

    # Crear directorio de figuras
    mkdir -p "$figures_dir"

    # Pandoc
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
    # shellcheck disable=SC2206
    [[ -n "$PANDOC_EXTRA_ARGS" ]] && pandoc_opts+=( $PANDOC_EXTRA_ARGS )

    local pandoc_ok=true
    if ! pandoc "${pandoc_opts[@]}" "$src_file" 2>/tmp/doc2md_stderr.txt; then
        pandoc_ok=false
    fi

    # Fallback LibreOffice para pptx/odp
    if ! $pandoc_ok && [[ "$src_ext" =~ ^(pptx|odp)$ ]]; then
        log_warn "Pandoc fallo, intentando pre-conversion con LibreOffice..."
        if libreoffice --headless --convert-to odt \
               --outdir "$tmp_media" "$src_file" &>/dev/null; then
            local lo_out
            lo_out=$(find "$tmp_media" -maxdepth 1 -name "*.odt" | head -1)
            if [[ -n "$lo_out" ]]; then
                pandoc_opts[0]="--from=odt"
                pandoc "${pandoc_opts[@]}" "$lo_out" 2>/tmp/doc2md_stderr.txt && \
                    pandoc_ok=true || true
            fi
        fi
    fi

    if ! $pandoc_ok; then
        log_err "Fallo la conversion: $src_file"
        log_err "  $(cat /tmp/doc2md_stderr.txt)"
        ERRORS+=("$src_file")
        COUNT_ERR=$(( COUNT_ERR + 1 ))
        rm -rf "$tmp_media"
        rmdir "$figures_dir" 2>/dev/null || true
        rmdir "${src_dir}/${name_no_ext}_files" 2>/dev/null || true
        return 1
    fi

    # Mover imagenes al directorio de figuras
    if [[ -d "$tmp_media" ]]; then
        find "$tmp_media" -type f \
            \( -iname "*.png" -o -iname "*.jpg" -o -iname "*.jpeg" \
               -o -iname "*.gif" -o -iname "*.svg" -o -iname "*.webp" \
               -o -iname "*.bmp" -o -iname "*.tiff" \) \
            -exec mv -n {} "$figures_dir/" \; 2>/dev/null || true
        rm -rf "$tmp_media"
    fi

    # Renombrar figuras
    local n_figs=0
    if [[ -d "$figures_dir" ]] && [[ -n "$(ls -A "$figures_dir" 2>/dev/null)" ]]; then
        n_figs=$(rename_figures "$figures_dir")
        if [[ "$IMG_FORMAT" != "png" ]]; then
            find "$figures_dir" -type f | while read -r img; do
                optimize_image "$img"
            done
        fi
    fi

    # Limpiar directorio de figuras si quedo vacio
    if [[ -d "$figures_dir" ]] && [[ -z "$(ls -A "$figures_dir" 2>/dev/null)" ]]; then
        rmdir "$figures_dir" 2>/dev/null || true
        rmdir "${src_dir}/${name_no_ext}_files" 2>/dev/null || true
    fi

    # Corregir rutas de imagen en el .md
    fix_image_paths "$out_md" "$figures_rel"

    COUNT_OK=$(( COUNT_OK + 1 ))
    CONVERTED_FILES+=("$src_file")
    log_ok "Listo → ${BOLD}$(basename "$out_md")${RESET}  [figuras: ${n_figs}]"
    return 0
}

# ---------------------------------------------------------------------------
# ELIMINAR ARCHIVOS CONVERTIDOS (solo los que tuvieron exito)
# ---------------------------------------------------------------------------
delete_converted() {
    [[ ${#CONVERTED_FILES[@]} -eq 0 ]] && return 0

    log_sep
    log "Eliminando archivos Office convertidos..."

    for src_file in "${CONVERTED_FILES[@]}"; do
        if $DRY_RUN; then
            log "[DRY-RUN] Se eliminaria: $src_file"
            COUNT_DELETE=$(( COUNT_DELETE + 1 ))
        elif [[ -f "$src_file" ]]; then
            rm -f "$src_file"
            COUNT_DELETE=$(( COUNT_DELETE + 1 ))
            log_ok "Eliminado: $(basename "$src_file")   ← $(dirname "$src_file")/"
        else
            log_warn "No encontrado para eliminar: $src_file"
        fi
    done
}

# ---------------------------------------------------------------------------
# ESCANEAR DIRECTORIO y convertir
# ---------------------------------------------------------------------------
run_conversions() {
    # Construir la expresion -name para find
    local name_exprs=()
    for fmt in "${FORMATS[@]}"; do
        name_exprs+=( -iname "*.${fmt}" -o )
    done
    unset 'name_exprs[${#name_exprs[@]}-1]'   # quitar el ultimo -o

    log "Directorio : $TARGET_DIR"
    log "Formatos   : ${FORMATS[*]}"
    if [[ ${#EXCLUDE_DIRS[@]} -gt 0 ]]; then
        log "Excluidas  : ${EXCLUDE_DIRS[*]}"
    fi
    $DRY_RUN         && log_warn "MODO DRY-RUN activo — no se escribira nada"
    $DELETE_CONVERTED && log_warn "DELETE-CONVERTED activo — se eliminaran originales convertidos"
    log_sep

    # Construir comando find con exclusiones como -prune
    local find_cmd=( find "$TARGET_DIR" )

    if [[ ${#EXCLUDE_DIRS[@]} -gt 0 ]]; then
        find_cmd+=( \( )
        local first=true
        for excl in "${EXCLUDE_DIRS[@]}"; do
            [[ -z "$excl" ]] && continue
            $first || find_cmd+=( -o )
            if [[ "$excl" == /* ]]; then
                find_cmd+=( -path "$excl" )
            else
                find_cmd+=( -name "$excl" )
            fi
            first=false
        done
        find_cmd+=( \) -prune -o )
    fi

    find_cmd+=( -type f \( "${name_exprs[@]}" \) -print0 )

    while IFS= read -r -d '' src; do
        convert_file "$src" || true
    done < <("${find_cmd[@]}" | sort -z)
}

# ---------------------------------------------------------------------------
# RESUMEN FINAL
# ---------------------------------------------------------------------------
print_summary() {
    local ts
    ts=$(date '+%Y-%m-%d %H:%M:%S')

    log_sep
    echo -e "${BOLD}RESUMEN${RESET}  [$ts]"
    echo -e "  Directorio            : $TARGET_DIR"
    echo -e "  Total encontrados     : $COUNT_TOTAL"
    echo -e "  ${GREEN}Convertidos OK        : $COUNT_OK${RESET}"
    echo -e "  ${YELLOW}Omitidos (ya existen) : $COUNT_SKIP${RESET}"
    echo -e "  ${RED}Con errores           : $COUNT_ERR${RESET}"
    $DELETE_CONVERTED && \
        echo -e "  ${RED}Originales eliminados : $COUNT_DELETE${RESET}"

    if [[ ${#ERRORS[@]} -gt 0 ]]; then
        echo -e "\n${RED}Archivos con error:${RESET}"
        for e in "${ERRORS[@]}"; do echo "  - $e"; done
    fi
    log_sep

    if [[ -n "$SUMMARY_FILE" ]]; then
        {
            echo "=== doc2md v2.0.0 — Resumen de conversion ==="
            echo "Fecha     : $ts"
            echo "Script    : ${SCRIPT_DIR}/${SCRIPT_NAME}"
            echo "Directorio: $TARGET_DIR"
            echo "Formatos  : ${FORMATS[*]}"
            echo "Excluidas : ${EXCLUDE_DIRS[*]:-ninguna}"
            echo "Total     : $COUNT_TOTAL"
            echo "OK        : $COUNT_OK"
            echo "Omitidos  : $COUNT_SKIP"
            echo "Errores   : $COUNT_ERR"
            $DELETE_CONVERTED && echo "Eliminados: $COUNT_DELETE"
            if [[ ${#ERRORS[@]} -gt 0 ]]; then
                echo ""
                echo "Archivos con error:"
                for e in "${ERRORS[@]}"; do echo "  $e"; done
            fi
            if [[ ${#CONVERTED_FILES[@]} -gt 0 ]]; then
                echo ""
                echo "Archivos convertidos:"
                for f in "${CONVERTED_FILES[@]}"; do echo "  $f"; done
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
    echo -e "${RESET}${BOLD}  Conversor recursivo Office → Markdown  v2.0.0${RESET}"
    echo -e "  ${CYAN}${SCRIPT_DIR}/${SCRIPT_NAME}${RESET}"
    echo ""

    parse_args "$@"
    check_deps
    run_conversions
    $DELETE_CONVERTED && delete_converted
    print_summary

    [[ $COUNT_ERR -gt 0 ]] && exit 1 || exit 0
}

main "$@"
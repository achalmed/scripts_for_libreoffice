#!/usr/bin/env bash
# lib/core/config.sh — Carga de configuración en tres capas.
# Precedencia: CLI > ~/.config/docflow/config.toml > config/defaults.conf.
# El TOML se parsea con tomllib (stdlib de Python ≥3.11) para no reinventar
# un parser frágil en Bash.

# config_load_defaults — carga config/defaults.conf y luego el TOML del usuario
config_load_defaults() {
    # shellcheck disable=SC1091
    source "${DOCFLOW_ROOT}/config/defaults.conf"

    local user_config="${DOCFLOW_CONFIG_FILE:-$DOCFLOW_CONFIG_FILE_DEFAULT}"
    [[ -f "$user_config" ]] && config_load_toml "$user_config"
    return 0
}

# config_load_toml <ruta> — traduce claves TOML a variables DOCFLOW_*
# La sección se convierte en prefijo: [markdown] flavor="gfm" → DOCFLOW_MD_FLAVOR
config_load_toml() {
    local toml_file="$1" lines
    lines="$(python3 - "$toml_file" <<'PYEOF'
import sys, tomllib

# Mapeo sección.clave → variable de entorno de docflow.
# Mantenerlo aquí (y no disperso) hace trivial añadir nuevas opciones.
KEYMAP = {
    "general.jobs": "DOCFLOW_JOBS",
    "general.overwrite": "DOCFLOW_OVERWRITE",
    "general.verify": "DOCFLOW_VERIFY",
    "general.cache": "DOCFLOW_CACHE",
    "general.output_dir": "DOCFLOW_OUTPUT_DIR",
    "general.default_excludes": "DOCFLOW_DEFAULT_EXCLUDES",
    "originals.delete": "DOCFLOW_DELETE_ORIGINALS",
    "originals.backup": "DOCFLOW_BACKUP_ORIGINALS",
    "originals.backup_dir_name": "DOCFLOW_BACKUP_DIR_NAME",
    "markdown.flavor": "DOCFLOW_MD_FLAVOR",
    "markdown.metadata": "DOCFLOW_MD_METADATA",
    "markdown.clean": "DOCFLOW_MD_CLEAN",
    "markdown.slide_separator": "DOCFLOW_MD_SLIDE_SEPARATOR",
    "markdown.include_notes": "DOCFLOW_MD_INCLUDE_NOTES",
    "markdown.pandoc_args": "DOCFLOW_PANDOC_ARGS",
    "images.format": "DOCFLOW_IMG_FORMAT",
    "images.quality": "DOCFLOW_IMG_QUALITY",
    "images.max_width": "DOCFLOW_IMG_MAX_WIDTH",
    "images.optimize": "DOCFLOW_IMG_OPTIMIZE",
    "pdf.engine": "DOCFLOW_PDF_ENGINE",
    "pdf.pdfa": "DOCFLOW_PDF_PDFA",
    "pdf.compress": "DOCFLOW_PDF_COMPRESS",
    "ocr.enabled": "DOCFLOW_OCR",
    "ocr.lang": "DOCFLOW_OCR_LANG",
    "watch.interval": "DOCFLOW_WATCH_INTERVAL",
    "watch.settle": "DOCFLOW_WATCH_SETTLE",
    "reports.formats": "DOCFLOW_REPORT_FORMATS",
    "reports.output_dir": "DOCFLOW_REPORT_OUT",
    "hooks.pre": "DOCFLOW_PRE_HOOK",
    "hooks.post": "DOCFLOW_POST_HOOK",
}

def flatten(d, prefix=""):
    for k, v in d.items():
        key = f"{prefix}{k}"
        if isinstance(v, dict):
            yield from flatten(v, f"{key}.")
        else:
            yield key, v

try:
    with open(sys.argv[1], "rb") as fh:
        data = tomllib.load(fh)
except Exception as exc:
    print(f"__TOML_ERROR__ {exc}")
    sys.exit(0)

import shlex
for key, value in flatten(data):
    var = KEYMAP.get(key)
    if var is None:
        print(f"__TOML_UNKNOWN__ {key}")
        continue
    if isinstance(value, bool):
        value = "true" if value else "false"
    elif isinstance(value, list):
        value = " ".join(str(x) for x in value)
    print(f"{var}={shlex.quote(str(value))}")
PYEOF
)" || return 0

    local line
    while IFS= read -r line; do
        [[ -z "$line" ]] && continue
        case "$line" in
            "__TOML_ERROR__"*)
                log_error "config.toml inválido: ${line#__TOML_ERROR__ }"
                exit "$EXIT_BAD_CONFIG" ;;
            "__TOML_UNKNOWN__"*)
                log_warn "Opción desconocida en config.toml: ${line#__TOML_UNKNOWN__ }" ;;
            *)
                eval "$line" ;;
        esac
    done <<< "$lines"
    log_trace "Configuración cargada desde $toml_file"
}

# config_write_template <ruta> — genera un config.toml comentado de ejemplo
config_write_template() {
    local dest="$1"
    mkdir -p "$(dirname -- "$dest")"
    cat > "$dest" <<'EOF'
# ~/.config/docflow/config.toml — Configuración de usuario de docflow.
# Todas las opciones son opcionales; lo no definido usa los valores por defecto.
# Precedencia: flags de la CLI > este archivo > defaults del proyecto.

[general]
# jobs = 8                # workers en paralelo (defecto: nproc)
# overwrite = false
# verify = true           # verificar salidas tras convertir
# cache = true            # no reconvertir archivos sin cambios (SHA-256)
# output_dir = ""         # vacío = junto al original
# default_excludes = ".git node_modules _backup"

[originals]
# delete = false          # eliminar originales convertidos con éxito
# backup = false          # mover originales a _backup/
# backup_dir_name = "_backup"

[markdown]
# flavor = "gfm"          # gfm | obsidian | logseq | mkdocs | hugo | quarto
# metadata = true         # frontmatter YAML con metadatos del documento
# clean = true            # limpieza del Markdown generado
# slide_separator = "---"
# include_notes = true    # notas del presentador en presentaciones
# pandoc_args = ""

[images]
# format = "keep"         # keep | png | jpg | webp | avif
# quality = 90
# max_width = 0           # 0 = sin redimensionar
# optimize = false

[pdf]
# engine = "auto"         # auto | soffice | pandoc | wkhtmltopdf | weasyprint
# pdfa = false
# compress = false

[ocr]
# enabled = false
# lang = "spa+eng"

[watch]
# interval = 30
# settle = 1

[reports]
# formats = "html,json"   # html | csv | json | md (separados por coma)
# output_dir = ""

[hooks]
# pre = ""                # comando antes de cada conversión
# post = ""               # comando después (recibe DOCFLOW_INPUT/OUTPUT/STATUS)
EOF
}

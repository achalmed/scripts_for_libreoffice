#!/usr/bin/env bash
# lib/config.sh — Configuración centralizada y constantes globales
# Por qué existe: centralizar aquí toda configuración evita "magic strings"
# dispersos y hace que ajustar el comportamiento sea un cambio en un solo lugar.

# ---------------------------------------------------------------------------
# Ruta canónica del proyecto (siempre fija, nunca se mueve)
# ---------------------------------------------------------------------------
readonly SCRIPT_DIR="$(cd "$(dirname "$(readlink -f "${BASH_SOURCE[0]}")")" && cd .. && pwd)"
readonly LIB_DIR="${SCRIPT_DIR}/lib"
readonly LOGS_DIR="${SCRIPT_DIR}/logs"

# ---------------------------------------------------------------------------
# Metadatos
# ---------------------------------------------------------------------------
readonly PROJECT_NAME="docflow"
readonly VERSION="2.0.0"
readonly AUTHOR="Edison Achalma"

# ---------------------------------------------------------------------------
# Valores por defecto para flags CLI (sobreescribibles desde cli.sh)
# ---------------------------------------------------------------------------
VERBOSE=false
DRY_RUN=false
DELETE_MODE="ask"       # ask | never | backup
BACKUP_SUFFIX="_backup"
PARALLEL_JOBS="$(nproc 2>/dev/null || echo 2)"
IMG_FORMAT="png"
IMG_QUALITY=90
WATCH_INTERVAL=30       # segundos entre ciclos de watch
REPORT_FORMAT="html"    # html | csv | none

# ---------------------------------------------------------------------------
# Directorios de exclusión por defecto para doc2md
# ---------------------------------------------------------------------------
declare -a DEFAULT_EXCLUDE_DIRS=(
    ".git"
    ".svn"
    "node_modules"
    "__pycache__"
    ".trash"
    "Trash"
    ".Trash"
)

# ---------------------------------------------------------------------------
# Mapeo de extensiones → formatos de salida para conversión a ODF
# Formato: "extension_entrada:extension_salida:descripción"
# ---------------------------------------------------------------------------
declare -a ODF_CONVERSION_MAP=(
    "docx:odt:Word Document"
    "doc:odt:Word 97-2003 Document"
    "dotx:ott:Word Template"
    "xlsx:ods:Excel Workbook"
    "xls:ods:Excel 97-2003 Workbook"
    "xlsm:ods:Excel Macro-Enabled Workbook"
    "xltx:ots:Excel Template"
    "pptx:odp:PowerPoint Presentation"
    "ppt:odp:PowerPoint 97-2003 Presentation"
    "pptm:odp:PowerPoint Macro-Enabled Presentation"
    "ppsx:odp:PowerPoint Show"
    "pps:odp:PowerPoint 97-2003 Show"
    "potx:otp:PowerPoint Template"
)

# Extensiones que se convierten directamente a PDF vía pandoc (sin LibreOffice)
declare -a PANDOC_TO_PDF_EXTS=("docx" "odt" "doc")

# Extensiones que requieren LibreOffice para convertir a PDF
declare -a SOFFICE_TO_PDF_EXTS=("pptx" "odp" "ppt" "pptm" "ppsx" "pps" "xlsx" "ods" "xls" "xlsm")

# ---------------------------------------------------------------------------
# Códigos de color para la terminal
# ---------------------------------------------------------------------------
if [[ -t 1 ]]; then
    readonly C_BOLD='\033[1m'
    readonly C_GREEN='\033[0;32m'
    readonly C_YELLOW='\033[1;33m'
    readonly C_RED='\033[0;31m'
    readonly C_BLUE='\033[0;34m'
    readonly C_CYAN='\033[0;36m'
    readonly C_RESET='\033[0m'
else
    # Sin colores si la salida no es una TTY (pipes, cron, logs)
    readonly C_BOLD='' C_GREEN='' C_YELLOW='' C_RED=''
    readonly C_BLUE='' C_CYAN='' C_RESET=''
fi

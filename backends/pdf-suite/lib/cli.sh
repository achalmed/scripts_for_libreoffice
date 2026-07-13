#!/usr/bin/env bash
# =============================================================================
# lib/cli.sh — Interfaz de línea de comandos y menú interactivo
# =============================================================================
# Dos modos de uso:
#   1. CLI directa : pdf-suite <operación> [opciones] <archivos>
#   2. Menú TUI    : pdf-suite (sin argumentos) → menú interactivo
# =============================================================================

[[ -n "${_CLI_SOURCED:-}" ]] && return 0
readonly _CLI_SOURCED=1

# show_banner()
# Imprime el banner del proyecto con versión y autor.
show_banner() {
    echo -e "${C_BOLD}${C_BLUE}"
    echo "  ╔══════════════════════════════════════════════════════════════╗"
    echo "  ║          PDF Suite v${PDF_SUITE_VERSION} — iLovePDF para Linux          ║"
    echo "  ║     Edison Achalma · UNSCH · Ayacucho, Perú                 ║"
    echo "  ╚══════════════════════════════════════════════════════════════╝"
    echo -e "${C_RESET}"
}

# show_help()
# Muestra la ayuda completa de la CLI.
show_help() {
    show_banner
    cat << EOF
${C_BOLD}USO${C_RESET}
  pdf-suite <operación> [opciones] <archivo(s)>
  pdf-suite                           → abre el menú interactivo

${C_BOLD}OPERACIONES${C_RESET}
  ${C_CYAN}compress${C_RESET}   Comprimir un PDF o carpeta de PDFs
  ${C_CYAN}merge${C_RESET}      Unir múltiples PDFs en uno
  ${C_CYAN}split${C_RESET}      Dividir un PDF por páginas o rangos
  ${C_CYAN}extract${C_RESET}    Extraer páginas específicas
  ${C_CYAN}rotate${C_RESET}     Rotar páginas
  ${C_CYAN}reorder${C_RESET}    Reordenar / invertir páginas
  ${C_CYAN}delete${C_RESET}     Eliminar páginas de un PDF
  ${C_CYAN}convert${C_RESET}    Convertir PDF ↔ imagen ↔ texto
  ${C_CYAN}ocr${C_RESET}        Aplicar OCR a PDFs escaneados
  ${C_CYAN}metadata${C_RESET}   Ver o editar metadatos XMP/DocInfo
  ${C_CYAN}protect${C_RESET}    Cifrar o descifrar un PDF
  ${C_CYAN}watermark${C_RESET}  Añadir marca de agua (texto o PDF)
  ${C_CYAN}repair${C_RESET}     Reparar o validar la estructura de un PDF
  ${C_CYAN}info${C_RESET}       Mostrar información detallada de un PDF
  ${C_CYAN}test${C_RESET}       Probar todos los métodos de compresión en un PDF
  ${C_CYAN}deps${C_RESET}       Verificar estado de dependencias instaladas

${C_BOLD}OPCIONES GLOBALES${C_RESET}
  ${C_YELLOW}-v, --verbose${C_RESET}         Mostrar detalles técnicos durante la ejecución
  ${C_YELLOW}-n, --dry-run${C_RESET}         Simular operaciones sin escribir ningún archivo
  ${C_YELLOW}-f, --force${C_RESET}           Sobreescribir archivos de salida existentes
  ${C_YELLOW}-r, --recursive${C_RESET}       Procesar subdirectorios (operaciones en carpeta)
  ${C_YELLOW}-o, --output PATH${C_RESET}     Ruta del archivo de salida
  ${C_YELLOW}-s, --suffix SUF${C_RESET}      Sufijo para archivos de salida (default: _out)
  ${C_YELLOW}    --log-file${C_RESET}        Guardar log en ~/.local/share/pdf-suite/logs/
  ${C_YELLOW}    --version${C_RESET}         Mostrar versión y salir
  ${C_YELLOW}-h, --help${C_RESET}            Mostrar esta ayuda

${C_BOLD}OPCIONES POR OPERACIÓN${C_RESET}
  compress:   -m METHOD  (screen|ebook|printer|prepress|ocr)
              -t PCT     umbral mínimo de reducción (default: 5)
  split:      --pages N  páginas por parte
  extract:    --pages RANGO  ej: "1-5", "1,3,7", "10-z"
  rotate:     --angle GRADOS  (90|180|270) [--pages RANGO]
  convert:    --to FORMAT  (png|jpg|txt|html|svg)
              --dpi N      resolución para imágenes (default: 150)
  ocr:        -l LANG     idioma Tesseract (default: spa)
  protect:    --encrypt   cifrar  |  --decrypt   descifrar
              --user-pass PASS   --owner-pass PASS
  watermark:  --text "TEXTO"  o  --stamp archivo.pdf
              --opacity 0.3  --angle 45  --color "0.8 0 0"
  metadata:   --set-title "T" --set-author "A" --set-subject "S"
              --set-keywords "k1,k2"

${C_BOLD}EJEMPLOS${C_RESET}
  pdf-suite compress -m ebook -r ~/Documents/biblioteca
  pdf-suite merge doc1.pdf doc2.pdf -o unido.pdf
  pdf-suite split --pages 10 libro.pdf
  pdf-suite extract --pages 1-5,28 informe.pdf
  pdf-suite ocr -l spa+eng escaneo.pdf
  pdf-suite metadata --set-author "Edison Achalma" *.pdf
  pdf-suite protect --encrypt --user-pass "1234" tesis.pdf
  pdf-suite watermark --text "BORRADOR" --opacity 0.3 doc.pdf
  pdf-suite convert --to png --dpi 300 presentacion.pdf
  pdf-suite test informe.pdf
EOF
}

# parse_global_flags()
# Extrae las flags globales del array de argumentos y ajusta las variables
# de configuración globales. Deja el resto de argumentos en REMAINING_ARGS.
#
# Arguments:
#   "$@" - todos los argumentos pasados al script
#
# Side effects:
#   Modifica VERBOSE, DRY_RUN, FORCE, RECURSIVE, DEFAULT_SUFFIX, LOG_TO_FILE
#   Popula REMAINING_ARGS con los argumentos no consumidos
parse_global_flags() {
    REMAINING_ARGS=()
    while [[ $# -gt 0 ]]; do
        case "$1" in
            -v|--verbose)    VERBOSE=true;                    shift ;;
            -n|--dry-run)    DRY_RUN=true;                   shift ;;
            -f|--force)      FORCE=true;                     shift ;;
            -r|--recursive)  RECURSIVE=true;                 shift ;;
            --log-file)      LOG_TO_FILE=true;               shift ;;
            -s|--suffix)     DEFAULT_SUFFIX="$2";            shift 2 ;;
            -o|--output)     OUTPUT_PATH="$2";               shift 2 ;;
            --version)       echo "pdf-suite v${PDF_SUITE_VERSION}"; exit 0 ;;
            -h|--help)       show_help; exit 0 ;;
            *)               REMAINING_ARGS+=("$1");         shift ;;
        esac
    done
}

# show_interactive_menu()
# Menú TUI de selección de operación cuando el usuario ejecuta sin argumentos.
# Usa select de Bash para una interfaz simple y portable.
show_interactive_menu() {
    show_banner
    echo -e "${C_BOLD}  Selecciona una operación:${C_RESET}\n"

    local options=(
        "compress  — Comprimir PDF(s)"
        "merge     — Unir PDFs"
        "split     — Dividir PDF por páginas"
        "extract   — Extraer páginas específicas"
        "rotate    — Rotar páginas"
        "reorder   — Reordenar / invertir páginas"
        "delete    — Eliminar páginas"
        "convert   — Convertir PDF ↔ imagen/texto"
        "ocr       — Aplicar OCR (PDFs escaneados)"
        "metadata  — Ver / editar metadatos"
        "protect   — Cifrar / descifrar"
        "watermark — Marca de agua"
        "repair    — Reparar / validar PDF"
        "info      — Información detallada"
        "test      — Comparar métodos de compresión"
        "deps      — Ver dependencias instaladas"
        "Salir"
    )

    PS3=$'\n  Tu elección: '
    select opt in "${options[@]}"; do
        local operation
        operation="$(echo "$opt" | awk '{print $1}')"
        if [[ "$operation" == "Salir" ]]; then
            echo -e "\n${C_DIM}Hasta luego.${C_RESET}"
            exit 0
        elif [[ -n "$operation" ]]; then
            echo ""
            INTERACTIVE_OPERATION="$operation"
            return 0
        else
            echo -e "${C_YELLOW}Opción no válida, intenta de nuevo.${C_RESET}"
        fi
    done
}

# prompt_for_file()
# Pide al usuario que ingrese la ruta de un archivo PDF de forma interactiva.
# Muestra las rutas de búsqueda conocidas como sugerencias.
#
# Outputs:
#   PROMPTED_FILE — path ingresado y validado
prompt_for_file() {
    local prompt_msg="${1:-Ingresa la ruta del PDF}"
    echo -e "\n${C_DIM}Rutas de búsqueda conocidas:${C_RESET}"
    for p in "${PDF_SEARCH_PATHS[@]}"; do
        [[ -d "$p" ]] && echo -e "  ${C_DIM}${p}${C_RESET}"
    done
    echo ""
    read -rp "  ${C_BOLD}${prompt_msg}: ${C_RESET}" PROMPTED_FILE
    PROMPTED_FILE="${PROMPTED_FILE/#\~/$HOME}"  # expande ~ manualmente
}

# prompt_for_option()
# Solicita un valor al usuario con un default visible.
#
# Arguments:
#   $1 - mensaje  $2 - valor default
#
# Outputs:
#   PROMPTED_VALUE — valor ingresado o el default si el usuario presiona Enter
prompt_for_option() {
    local message="$1"
    local default="$2"
    read -rp "  ${message} [${C_CYAN}${default}${C_RESET}]: " PROMPTED_VALUE
    PROMPTED_VALUE="${PROMPTED_VALUE:-$default}"
}

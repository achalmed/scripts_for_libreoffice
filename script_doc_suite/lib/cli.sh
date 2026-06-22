#!/usr/bin/env bash
# lib/cli.sh — Interfaz de línea de comandos unificada
# Por qué existe: centralizar todo el parsing aquí mantiene main.sh limpio
# y hace que agregar nuevos flags sea un cambio en un solo lugar.

# ---------------------------------------------------------------------------
# show_help()
# Muestra el mensaje de ayuda completo y sale con código 0.
# ---------------------------------------------------------------------------
show_help() {
    cat <<EOF

${C_BOLD}docflow v${VERSION}${C_RESET} — Suite de conversión de documentos Office
${C_BOLD}Autor:${C_RESET} ${AUTHOR}

${C_BOLD}USO:${C_RESET}
  docflow <COMANDO> [OPCIONES] <RUTA>

${C_BOLD}COMANDOS:${C_RESET}
  ${C_GREEN}to-odf${C_RESET}     Convierte archivos MS Office (.docx/.pptx/.xlsx...) → OpenDocument
  ${C_GREEN}to-pdf${C_RESET}     Convierte archivos Office/ODF directamente a PDF
  ${C_GREEN}to-md${C_RESET}      Convierte .docx/.odt/.pptx/.odp → Markdown (recursivo)
  ${C_GREEN}watch${C_RESET}      Vigila una carpeta y convierte automáticamente al detectar cambios
  ${C_GREEN}report${C_RESET}     Genera reporte HTML o CSV de la última sesión de conversión

${C_BOLD}OPCIONES GLOBALES:${C_RESET}
  -v, --verbose           Mostrar información detallada del proceso
  -n, --dry-run           Simular sin escribir ni eliminar nada
  -l, --log <archivo>     Guardar log en archivo además de stdout
  -h, --help              Mostrar esta ayuda
      --version           Mostrar versión

${C_BOLD}OPCIONES DE ENTRADA/SALIDA:${C_RESET}
  -i, --input <ruta>      Archivo o directorio a procesar
  -o, --output <dir>      Directorio de salida (por defecto: mismo que el origen)
  -f, --formats <lista>   Formatos a procesar, separados por coma
                          Ej: docx,pptx,xlsx

${C_BOLD}OPCIONES DE ARCHIVOS ORIGINALES:${C_RESET}
      --delete            Eliminar originales después de conversión exitosa
                          (pide confirmación antes de borrar)
      --backup            Mover originales a carpeta _backup/ en vez de eliminar
  -w, --overwrite         Sobreescribir archivos de destino ya existentes

${C_BOLD}OPCIONES DE EXCLUSIÓN (solo to-md):${C_RESET}
  -e, --exclude <nombre>  Excluir carpeta por nombre o ruta absoluta
                          Usar varias veces para múltiples exclusiones
      --no-default-excludes  No aplicar exclusiones por defecto

${C_BOLD}OPCIONES DE IMÁGENES (solo to-md):${C_RESET}
      --img-format <fmt>  Formato de imágenes: png | jpg | webp (defecto: png)
      --img-quality <n>   Calidad para jpg/webp, 1-100 (defecto: 90)

${C_BOLD}OPCIONES DE PANDOC (solo to-md):${C_RESET}
      --pandoc-args <args>  Argumentos extra para pandoc (entre comillas)

${C_BOLD}OPCIONES DE REPORTE:${C_RESET}
      --report-format <fmt>  Formato del reporte: html | csv (defecto: html)
      --report-out <archivo> Ruta del archivo de reporte a generar

${C_BOLD}OPCIONES DE WATCH:${C_RESET}
      --interval <seg>    Intervalo de revisión en segundos (defecto: 30)
      --watch-cmd <cmd>   Comando adicional a ejecutar post-conversión

${C_BOLD}EJEMPLOS:${C_RESET}
  # Convertir todos los .docx y .pptx de una carpeta a ODF
  docflow to-odf -i ~/Documents/trabajos -f docx,pptx

  # Convertir un archivo específico directamente a PDF
  docflow to-pdf -i ~/Documents/informe.docx -o ~/Escritorio/

  # Convertir carpeta ideas a Markdown, excluir borradores
  docflow to-md -i ~/Documents/ideas -e borradores -w

  # Vigilar carpeta y convertir al vuelo a PDF
  docflow watch -i ~/Descargas --formats docx,pptx

  # Generar reporte HTML de la última sesión
  docflow report --report-format html --report-out ~/logs/reporte.html

EOF
}

# ---------------------------------------------------------------------------
# parse_arguments()
# Parsea todos los flags de la línea de comandos.
# Popula variables globales que los módulos de lógica consumen.
# ---------------------------------------------------------------------------
parse_arguments() {
    # Requerir al menos un argumento
    if [[ $# -eq 0 ]]; then
        show_help
        exit 0
    fi

    # Primer argumento posicional = comando
    COMMAND="${1:-}"
    shift || true

    # Variables específicas de to-md
    declare -ga EXCLUDE_DIRS=("${DEFAULT_EXCLUDE_DIRS[@]}")
    USE_DEFAULT_EXCLUDES=true
    PANDOC_EXTRA_ARGS=""
    OVERWRITE=false

    # Variables de entrada/salida
    INPUT_PATH=""
    OUTPUT_DIR=""
    FORMATS_FILTER=""
    LOG_FILE_PATH=""
    REPORT_OUT=""
    WATCH_CMD=""
    DELETE_ORIGINALS=false
    BACKUP_ORIGINALS=false

    while [[ $# -gt 0 ]]; do
        case "$1" in
            -i|--input)
                INPUT_PATH="$2"; shift 2 ;;
            -o|--output)
                OUTPUT_DIR="$2"; shift 2 ;;
            -f|--formats)
                FORMATS_FILTER="$2"; shift 2 ;;
            -v|--verbose)
                VERBOSE=true; shift ;;
            -n|--dry-run)
                DRY_RUN=true; shift ;;
            -w|--overwrite)
                OVERWRITE=true; shift ;;
            --delete)
                DELETE_ORIGINALS=true; shift ;;
            --backup)
                BACKUP_ORIGINALS=true; shift ;;
            -l|--log)
                LOG_FILE_PATH="$2"; shift 2 ;;
            -e|--exclude)
                # Soporte de lista separada por coma o flag repetido
                IFS=',' read -ra _excl <<< "$2"
                EXCLUDE_DIRS+=("${_excl[@]}")
                shift 2 ;;
            --no-default-excludes)
                USE_DEFAULT_EXCLUDES=false
                EXCLUDE_DIRS=()
                shift ;;
            --img-format)
                IMG_FORMAT="$2"; shift 2 ;;
            --img-quality)
                IMG_QUALITY="$2"; shift 2 ;;
            --pandoc-args)
                PANDOC_EXTRA_ARGS="$2"; shift 2 ;;
            --interval)
                WATCH_INTERVAL="$2"; shift 2 ;;
            --watch-cmd)
                WATCH_CMD="$2"; shift 2 ;;
            --report-format)
                REPORT_FORMAT="$2"; shift 2 ;;
            --report-out)
                REPORT_OUT="$2"; shift 2 ;;
            -h|--help)
                show_help; exit 0 ;;
            --version)
                echo "${PROJECT_NAME} v${VERSION}"; exit 0 ;;
            -*)
                log_error "Flag desconocido: $1"
                echo "  Ejecuta 'docflow --help' para ver opciones disponibles."
                exit 2 ;;
            *)
                # Argumento posicional sin flag → atajar como --input
                if [[ -z "$INPUT_PATH" ]]; then
                    INPUT_PATH="$1"
                fi
                shift ;;
        esac
    done

    # Activar log a archivo si se especificó
    if [[ -n "$LOG_FILE_PATH" ]]; then
        setup_log_file "$LOG_FILE_PATH"
    fi

    # Si --backup y --delete se usan juntos, --backup tiene precedencia
    if [[ "$BACKUP_ORIGINALS" == "true" ]]; then
        DELETE_ORIGINALS=false
    fi
}

# ---------------------------------------------------------------------------
# validate_command()
# Verifica que el comando dado sea uno de los reconocidos.
# ---------------------------------------------------------------------------
validate_command() {
    case "${COMMAND:-}" in
        to-odf|to-pdf|to-md|watch|report) return 0 ;;
        "")
            log_error "Se requiere un comando."
            show_help; exit 2 ;;
        *)
            log_error "Comando desconocido: '${COMMAND}'"
            echo "  Comandos válidos: to-odf, to-pdf, to-md, watch, report"
            exit 2 ;;
    esac
}

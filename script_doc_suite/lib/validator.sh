#!/usr/bin/env bash
# lib/validator.sh — Validación de dependencias y entradas
# Por qué existe: centralizar todas las validaciones aquí garantiza que
# el script falle rápido y con mensajes claros antes de tocar ningún archivo.

# ---------------------------------------------------------------------------
# validate_dependencies()
# Verifica que las herramientas requeridas estén instaladas.
# Recibe una lista de comandos como argumentos.
# Retorna: 0 si todas presentes, 1 si alguna falta (con lista de ausentes)
# ---------------------------------------------------------------------------
validate_dependencies() {
    local -a missing=()
    for cmd in "$@"; do
        if ! command -v "$cmd" &>/dev/null; then
            missing+=("$cmd")
        fi
    done

    if [[ ${#missing[@]} -gt 0 ]]; then
        log_error "Dependencias faltantes: ${missing[*]}"
        echo ""
        echo "  Instala en Arch Linux:"
        for dep in "${missing[@]}"; do
            case "$dep" in
                pandoc)      echo "    sudo pacman -S pandoc" ;;
                python3)     echo "    sudo pacman -S python" ;;
                soffice)     echo "    sudo pacman -S libreoffice-still" ;;
                libreoffice) echo "    sudo pacman -S libreoffice-still" ;;
                convert)     echo "    sudo pacman -S imagemagick" ;;
                inotifywait) echo "    sudo pacman -S inotify-tools" ;;
                parallel)    echo "    sudo pacman -S parallel" ;;
                *)           echo "    sudo pacman -S ${dep}" ;;
            esac
        done
        echo ""
        return 1
    fi
    return 0
}

# ---------------------------------------------------------------------------
# validate_path_exists()
# Verifica que una ruta exista (archivo o directorio).
# Argumentos: $1=ruta, $2=tipo esperado (file|dir|any)
# ---------------------------------------------------------------------------
validate_path_exists() {
    local path="$1"
    local kind="${2:-any}"

    if [[ -z "$path" ]]; then
        log_error "Se requiere una ruta pero no se proporcionó ninguna."
        return 1
    fi

    case "$kind" in
        file)
            if [[ ! -f "$path" ]]; then
                log_error "No se encontró el archivo: $path"
                return 1
            fi
            ;;
        dir)
            if [[ ! -d "$path" ]]; then
                log_error "No se encontró el directorio: $path"
                return 1
            fi
            ;;
        any)
            if [[ ! -e "$path" ]]; then
                log_error "No existe la ruta: $path"
                return 1
            fi
            ;;
    esac
    return 0
}

# ---------------------------------------------------------------------------
# validate_writable_dir()
# Verifica que el directorio de destino sea escribible, creándolo si no existe.
# Argumentos: $1=ruta del directorio
# ---------------------------------------------------------------------------
validate_writable_dir() {
    local dir="$1"
    if [[ ! -d "$dir" ]]; then
        if ! mkdir -p "$dir" 2>/dev/null; then
            log_error "No se pudo crear el directorio: $dir"
            return 1
        fi
        log_debug "Directorio creado: $dir"
    fi
    if [[ ! -w "$dir" ]]; then
        log_error "Sin permisos de escritura en: $dir"
        return 1
    fi
    return 0
}

# ---------------------------------------------------------------------------
# validate_extension()
# Verifica que la extensión de un archivo esté en la lista de soportados.
# Argumentos: $1=archivo, $2...=extensiones válidas
# ---------------------------------------------------------------------------
validate_extension() {
    local file="$1"
    shift
    local ext="${file##*.}"
    ext="${ext,,}"  # lowercase

    for valid in "$@"; do
        if [[ "$ext" == "${valid,,}" ]]; then
            return 0
        fi
    done

    log_error "Extensión '.${ext}' no soportada para: $(basename "$file")"
    log_warn  "Extensiones válidas: $*"
    return 1
}

# ---------------------------------------------------------------------------
# resolve_output_path()
# Calcula la ruta de salida en función del modo.
# Argumentos: $1=archivo_entrada, $2=extension_salida, $3=outdir (opcional)
# Imprime la ruta calculada en stdout.
# ---------------------------------------------------------------------------
resolve_output_path() {
    local input_file="$1"
    local out_ext="$2"
    local out_dir="${3:-}"

    local base_name
    base_name="$(basename "${input_file%.*}")"

    if [[ -n "$out_dir" ]]; then
        echo "${out_dir}/${base_name}.${out_ext}"
    else
        # Mismo directorio que el archivo de entrada
        echo "$(dirname "$input_file")/${base_name}.${out_ext}"
    fi
}

# ---------------------------------------------------------------------------
# confirm_action()
# Pide confirmación interactiva al usuario. Retorna 0 si confirma, 1 si no.
# Argumentos: $1=mensaje de pregunta
# ---------------------------------------------------------------------------
confirm_action() {
    local message="$1"
    local response
    echo -e "\n${C_YELLOW}${message}${C_RESET}"
    read -rp "  [s/N] → " response
    [[ "${response,,}" == "s" ]]
}

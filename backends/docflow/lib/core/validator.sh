#!/usr/bin/env bash
# lib/core/validator.sh — Validaciones de entrada. Falla rápido y claro
# antes de tocar cualquier archivo.

# validate_input_path <ruta> → 0 si es archivo o directorio legible
validate_input_path() {
    local path="${1:-}"
    if [[ -z "$path" ]]; then
        log_error "Se requiere una ruta de entrada."
        return 1
    fi
    if [[ ! -e "$path" ]]; then
        log_error "La ruta no existe: $path"
        return 1
    fi
    if [[ ! -r "$path" ]]; then
        log_error "Sin permiso de lectura: $path"
        return 1
    fi
    return 0
}

# validate_writable_dir <dir> — crea el directorio si no existe y verifica escritura
validate_writable_dir() {
    local dir="$1"
    mkdir -p -- "$dir" 2>/dev/null || { log_error "No se pudo crear: $dir"; return 1; }
    [[ -w "$dir" ]] || { log_error "Sin permiso de escritura: $dir"; return 1; }
    return 0
}

# validate_protected <archivo> → 0 si el documento parece protegido/corrupto.
# Los OOXML/ODF cifrados con contraseña no son contenedores ZIP (son OLE),
# lo que permite detectarlos sin abrirlos.
validate_protected() {
    local file="$1" ext
    ext="$(util_ext "$file")"
    case "$ext" in
        docx|docm|dotx|xlsx|xlsm|xltx|pptx|pptm|ppsx|potx|odt|ott|odp|otp|ods|ots|epub)
            util_is_zip "$file" && return 1
            return 0 ;;
        pdf)
            deps_has qpdf || return 1
            [[ "$(qpdf --is-encrypted "$file" 2>/dev/null; echo $?)" == "0" ]] && return 0
            return 1 ;;
        *) return 1 ;;
    esac
}

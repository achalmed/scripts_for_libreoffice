#!/usr/bin/env bash
# lib/core/utils.sh — Utilidades puras y reutilizables.
# Regla: nada aquí puede depender de estado global de docflow ni hacer I/O
# destructivo; solo funciones auxiliares deterministas.

# util_ext <ruta> → extensión en minúsculas, sin punto ("" si no tiene)
util_ext() {
    local base="${1##*/}"
    [[ "$base" == *.* ]] || { printf ''; return 0; }
    local ext="${base##*.}"
    printf '%s' "${ext,,}"
}

# util_basename_noext <ruta> → nombre base sin extensión
util_basename_noext() {
    local base="${1##*/}"
    printf '%s' "${base%.*}"
}

# util_in_list <valor> <item>... → 0 si el valor está en la lista
util_in_list() {
    local target="$1" item
    shift
    for item in "$@"; do [[ "$item" == "$target" ]] && return 0; done
    return 1
}

# util_human_size <bytes> → tamaño legible (1.4 MiB)
util_human_size() {
    local b="${1:-0}"
    if   (( b >= 1073741824 )); then printf '%d.%d GiB' $((b/1073741824)) $((b%1073741824*10/1073741824))
    elif (( b >= 1048576 ));    then printf '%d.%d MiB' $((b/1048576)) $((b%1048576*10/1048576))
    elif (( b >= 1024 ));       then printf '%d.%d KiB' $((b/1024)) $((b%1024*10/1024))
    else printf '%d B' "$b"; fi
}

# util_file_size <ruta> → bytes (0 si no existe)
util_file_size() {
    [[ -f "$1" ]] && stat -c%s -- "$1" 2>/dev/null || printf '0'
}

# util_epoch_ms → milisegundos desde epoch (para perfilado)
util_epoch_ms() {
    printf '%d' "$(( $(date +%s%N) / 1000000 ))"
}

# util_sha256 <ruta> → hash SHA-256 del contenido
util_sha256() {
    sha256sum -- "$1" 2>/dev/null | cut -d' ' -f1
}

# util_mktemp_dir <etiqueta> → directorio temporal propio de docflow
util_mktemp_dir() {
    mktemp -d "${TMPDIR:-/tmp}/docflow_${1:-tmp}_XXXXXX"
}

# util_confirm <pregunta> → 0 si el usuario confirma (o --force activo)
util_confirm() {
    [[ "${DOCFLOW_FORCE:-false}" == "true" ]] && return 0
    # En entornos no interactivos sin --force, la respuesta segura es "no".
    [[ -t 0 ]] || return 1
    local reply
    read -r -p "${1} [s/N] " reply
    [[ "$reply" =~ ^([sS]|[sS][iI]|[yY])$ ]]
}

# util_relpath <destino> <base> → ruta relativa de destino respecto a base
util_relpath() {
    python3 -c 'import os,sys; print(os.path.relpath(sys.argv[1], sys.argv[2]))' "$1" "$2"
}

# util_is_zip <ruta> → 0 si el archivo es un contenedor ZIP válido.
# Un .docx/.pptx/.odt que NO es zip suele estar protegido con contraseña
# (contenedor OLE cifrado) o corrupto.
util_is_zip() {
    python3 -c 'import sys,zipfile; sys.exit(0 if zipfile.is_zipfile(sys.argv[1]) else 1)' "$1" 2>/dev/null
}

# util_join <separador> <items...> → une items con separador
util_join() {
    local sep="$1" out="" item
    shift
    for item in "$@"; do out+="${out:+$sep}${item}"; done
    printf '%s' "$out"
}

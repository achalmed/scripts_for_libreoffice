#!/usr/bin/env bash
# lib/core/soffice.sh — Wrapper robusto de LibreOffice headless.
# Por qué existe: soffice tiene dos problemas serios en automatización:
#   1. Dos instancias simultáneas corrompen el perfil de usuario compartido
#      → cada invocación usa un perfil aislado (-env:UserInstallation).
#   2. Puede colgarse indefinidamente con documentos rotos → timeout.

DOCFLOW_SOFFICE_TIMEOUT="${DOCFLOW_SOFFICE_TIMEOUT:-300}"

# soffice_convert <entrada> <filtro_destino> <dir_salida>
# filtro_destino: extensión ("pdf") o extensión:Filtro ("pdf:impress_pdf_Export")
# Devuelve 0 y garantiza que el archivo esperado existe; ≠0 en caso contrario.
soffice_convert() {
    local input="$1" target="$2" out_dir="$3"
    local profile rc output_name
    profile="$(util_mktemp_dir lo_profile)"

    # Nota: soffice no soporta el separador "--"; se pasa la ruta tal cual.
    log_trace "soffice: $input → $target (perfil: $profile)"
    timeout "$DOCFLOW_SOFFICE_TIMEOUT" soffice --headless --norestore \
        "-env:UserInstallation=file://${profile}" \
        --convert-to "$target" --outdir "$out_dir" "$input" \
        >/dev/null 2>&1
    rc=$?
    rm -rf -- "$profile"

    if (( rc == 124 )); then
        log_error "LibreOffice superó el timeout (${DOCFLOW_SOFFICE_TIMEOUT}s): $(basename -- "$input")"
        return 1
    fi

    # soffice puede devolver 0 sin producir salida (p. ej. documento protegido
    # por contraseña) → la única verificación fiable es que el archivo exista.
    output_name="$(util_basename_noext "$input").${target%%:*}"
    [[ -s "${out_dir}/${output_name}" ]]
}

# soffice_intermediate <entrada> <ext_destino>
# Conversión a formato intermedio en un directorio temporal.
# Imprime la ruta del archivo generado; el llamador debe borrar el directorio.
soffice_intermediate() {
    local input="$1" target_ext="$2" tmp_dir
    tmp_dir="$(util_mktemp_dir intermediate)"
    if soffice_convert "$input" "$target_ext" "$tmp_dir"; then
        printf '%s' "${tmp_dir}/$(util_basename_noext "$input").${target_ext}"
        return 0
    fi
    rm -rf -- "$tmp_dir"
    return 1
}

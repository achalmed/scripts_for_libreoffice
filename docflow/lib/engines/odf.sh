#!/usr/bin/env bash
# lib/engines/odf.sh — ODF Engine: MS Office → OpenDocument.
# El mapa de destinos vive en el registry (lib/formats/*): docx→odt,
# xlsx→ods, pptx→odp, y plantillas dotx→ott, xltx→ots, potx→otp.

engine_odf_output_ext() {
    printf '%s' "${REG_ODF_TARGET[$1]:-}"
}

engine_odf_convert() {
    local input="$1" output="$2"
    deps_require soffice
    local out_dir produced
    out_dir="$(dirname -- "$output")"
    soffice_convert "$input" "$(util_ext "$output")" "$out_dir" || return 1
    produced="${out_dir}/$(util_basename_noext "$input").$(util_ext "$output")"
    [[ "$produced" != "$output" && -f "$produced" ]] && mv -- "$produced" "$output"
    [[ -s "$output" ]]
}

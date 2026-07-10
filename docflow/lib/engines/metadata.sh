#!/usr/bin/env bash
# lib/engines/metadata.sh — Metadata Engine.
# Extrae metadatos del documento original (título, autor, fechas, idioma,
# palabras clave, empresa, comentarios) leyendo directamente el XML de los
# contenedores OOXML/ODF — sin dependencias externas — y los inyecta como
# frontmatter YAML en el Markdown generado.

# metadata_yaml <archivo> → bloque YAML por stdout (vacío si no hay metadatos)
metadata_yaml() {
    python3 "${DOCFLOW_ROOT}/lib/helpers/office_meta.py" "$1" 2>/dev/null
}

# metadata_prepend_frontmatter <original> <md>
metadata_prepend_frontmatter() {
    local original="$1" md_file="$2"
    local yaml
    yaml="$(metadata_yaml "$original")"
    [[ -z "$yaml" ]] && return 0

    local tmp
    tmp="$(mktemp "${TMPDIR:-/tmp}/docflow_fm_XXXXXX")"
    {
        printf -- '---\n%s\n---\n\n' "$yaml"
        cat -- "$md_file"
    } > "$tmp" && mv -- "$tmp" "$md_file"
}

# metadata_show <archivo> — para `docflow pdf info` y usos interactivos
metadata_show() {
    local yaml
    yaml="$(metadata_yaml "$1")"
    if [[ -n "$yaml" ]]; then
        log_plain "$yaml"
    else
        log_info "Sin metadatos legibles: $(basename -- "$1")"
    fi
}

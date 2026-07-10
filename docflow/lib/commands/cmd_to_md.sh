#!/usr/bin/env bash
# lib/commands/cmd_to_md.sh — Comando to-md: cualquier formato → Markdown.
# Envoltorio fino: todo el trabajo real ocurre en dispatch + Markdown Engine.

cmd_to_md() {
    deps_require pandoc python3
    log_header "Conversión → Markdown"
    if [[ $# -eq 0 ]]; then
        log_error "Se requiere al menos una ruta."
        log_plain "  Uso: docflow to-md [opciones] <archivo|directorio>..."
        exit "$EXIT_USAGE"
    fi
    dispatch_run md "$@"
    exit $?
}

cli_help_to_md() {
    cat >&2 <<EOF
${C_BOLD}docflow to-md${C_RESET} — Convierte documentos a Markdown de alta calidad.

USO: docflow to-md [opciones] <archivo|directorio>...

Formatos: $(registry_exts_for md | paste -sd' ')

OPCIONES ESPECÍFICAS:
  --flavor <f>          gfm | obsidian | logseq | mkdocs | hugo | quarto
  --no-metadata         No añadir frontmatter YAML con metadatos
  --no-clean            No limpiar el Markdown generado
  --no-notes            No exportar notas del presentador
  --slide-separator <s> Separador entre diapositivas (defecto: ---)
  --img-format <fmt>    keep | png | jpg | webp | avif
  --img-quality <n>     Calidad 1-100 (defecto: 90)
  --img-max-width <px>  Redimensionar imágenes anchas (0 = no)
  --img-optimize        Optimizar imágenes extraídas
  --ocr                 OCR para PDFs escaneados (requiere ocrmypdf)
  --pandoc-args <args>  Argumentos extra para pandoc

EJEMPLOS:
  docflow to-md tesis/capitulo1.docx
  docflow to-md --flavor obsidian --img-format webp ~/Documentos/apuntes/
  docflow to-md -f pptx,odp --no-notes ~/Documentos/clases/
EOF
}

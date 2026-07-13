# Autocompletado bash para docflow
_docflow() {
    local cur prev commands
    COMPREPLY=()
    cur="${COMP_WORDS[COMP_CWORD]}"
    prev="${COMP_WORDS[COMP_CWORD-1]}"
    commands="to-md to-pdf to-odf to-office pdf watch report index doctor formats cache config help version"

    if [[ $COMP_CWORD -eq 1 ]]; then
        mapfile -t COMPREPLY < <(compgen -W "$commands" -- "$cur")
        return 0
    fi

    case "${COMP_WORDS[1]}" in
        pdf)
            if [[ $COMP_CWORD -eq 2 ]]; then
                mapfile -t COMPREPLY < <(compgen -W "merge split rotate extract compress pdfa encrypt decrypt watermark ocr info" -- "$cur")
                return 0
            fi ;;
        cache)
            mapfile -t COMPREPLY < <(compgen -W "stats clear prune" -- "$cur"); return 0 ;;
        config)
            mapfile -t COMPREPLY < <(compgen -W "init show path" -- "$cur"); return 0 ;;
    esac

    case "$prev" in
        --flavor)     mapfile -t COMPREPLY < <(compgen -W "gfm obsidian logseq mkdocs hugo quarto" -- "$cur"); return 0 ;;
        --img-format) mapfile -t COMPREPLY < <(compgen -W "keep png jpg webp avif" -- "$cur"); return 0 ;;
        --pdf-engine) mapfile -t COMPREPLY < <(compgen -W "auto soffice pandoc wkhtmltopdf weasyprint" -- "$cur"); return 0 ;;
        --report)     mapfile -t COMPREPLY < <(compgen -W "html csv json md" -- "$cur"); return 0 ;;
        --log-level)  mapfile -t COMPREPLY < <(compgen -W "trace debug info warn error quiet" -- "$cur"); return 0 ;;
    esac

    if [[ "$cur" == -* ]]; then
        mapfile -t COMPREPLY < <(compgen -W "-o --output-dir -f --formats -e --exclude --exclude-regex --include \
            -j --jobs -w --overwrite -n --dry-run --resume --no-cache --no-verify --force \
            --backup --delete -v --verbose -q --quiet --json -l --log --report --report-out \
            --flavor --no-metadata --no-clean --no-notes --img-format --img-quality \
            --img-max-width --img-optimize --pdf-engine --pdfa --compress --ocr \
            --pre-hook --post-hook --config --help" -- "$cur")
        return 0
    fi

    # Por defecto: rutas
    COMPREPLY=()
    return 0
}
complete -o default -F _docflow docflow

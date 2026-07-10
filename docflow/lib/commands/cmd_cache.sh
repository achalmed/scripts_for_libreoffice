#!/usr/bin/env bash
# lib/commands/cmd_cache.sh — Comando cache: gestión del caché de conversiones.

cmd_cache() {
    case "${1:-stats}" in
        stats) cache_stats ;;
        clear) cache_clear ;;
        prune) cache_prune ;;
        *)
            log_error "Uso: docflow cache [stats|clear|prune]"
            exit "$EXIT_USAGE" ;;
    esac
    exit 0
}

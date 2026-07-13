#!/usr/bin/env bash
# lib/commands/cmd___worker.sh — Subcomando interno para paralelización.
# NO es parte de la CLI pública. parallel.sh lanza:
#     docflow __worker <task> <archivo>
# El worker hereda la configuración por variables DOCFLOW_* exportadas y
# registra su resultado en el results.tsv de la sesión compartida.

cmd___worker() {
    local task="$1" file="$2"
    # Los workers no dibujan progreso ni cabeceras; solo convierten.
    dispatch_convert_one "$task" "$file"
    exit 0
}

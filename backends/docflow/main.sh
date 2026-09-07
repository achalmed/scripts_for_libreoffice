#!/usr/bin/env bash
# main.sh — docflow: entrada según el patrón main + config + lib (FS3, 2026-09-07). Delega en bin/docflow, que
# carga lib/core/bootstrap.sh y config/defaults.conf; el logger propio (niveles, quiet) se mantiene por diseño.
exec "$(cd "$(dirname "$(readlink -f "${BASH_SOURCE[0]}")")" && pwd)/bin/docflow" "$@"

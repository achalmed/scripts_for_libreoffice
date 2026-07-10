#!/usr/bin/env bash
# tests/test_helper.bash — Utilidades comunes de la suite Bats.

DOCFLOW_ROOT="$(cd "$(dirname "$(readlink -f "${BASH_SOURCE[0]}")")/.." && pwd)"
export DOCFLOW_ROOT
DOCFLOW_BIN="${DOCFLOW_ROOT}/bin/docflow"
export DOCFLOW_BIN

# Cada test corre en un sandbox aislado (config, caché y estado propios)
setup_sandbox() {
    TEST_DIR="$(mktemp -d "${TMPDIR:-/tmp}/docflow_test_XXXXXX")"
    export TEST_DIR
    export XDG_CONFIG_HOME="${TEST_DIR}/config"
    export XDG_CACHE_HOME="${TEST_DIR}/cache"
    export XDG_STATE_HOME="${TEST_DIR}/state"
    export NO_COLOR=1
    cd "$TEST_DIR" || exit 1
}

teardown_sandbox() {
    cd / && rm -rf -- "$TEST_DIR"
}

# Genera un docx de prueba con pandoc (fixture mínima y determinista)
make_docx() {
    local name="${1:-prueba.docx}"
    cat > "${TEST_DIR}/_src.md" <<'EOF'
# Título de prueba

Un párrafo con **negrita** y una lista:

- uno
- dos

| A | B |
|---|---|
| 1 | 2 |
EOF
    pandoc "${TEST_DIR}/_src.md" -o "${TEST_DIR}/${name}"
    rm -f "${TEST_DIR}/_src.md"
}

have() { command -v "$1" &>/dev/null; }

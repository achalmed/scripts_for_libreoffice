#!/usr/bin/env bash
# install.sh — Instalador de DocFlow Studio
#
# 1. Crea un entorno virtual local (.venv) e instala PySide6 + dependencias.
# 2. Enlaza el lanzador en ~/.local/bin/docflow-studio.
# 3. Instala la entrada de escritorio (menú de aplicaciones).
#
# Los backends Bash (docflow, pdf-suite) tienen sus propios instaladores en
# backends/*/install.sh para dependencias del sistema (pandoc, qpdf, gs…);
# este script solo prepara el frontend.
set -euo pipefail

DIR="$(cd "$(dirname "$(readlink -f "${BASH_SOURCE[0]}")")" && pwd)"
BIN_DIR="${HOME}/.local/bin"
APPS_DIR="${HOME}/.local/share/applications"

echo "==> Creando entorno virtual (.venv)…"
python3 -m venv "${DIR}/.venv"
"${DIR}/.venv/bin/pip" install --upgrade pip >/dev/null
"${DIR}/.venv/bin/pip" install -r "${DIR}/requirements.txt"

echo "==> Enlazando lanzador en ${BIN_DIR}/docflow-studio…"
mkdir -p "$BIN_DIR"
chmod +x "${DIR}/run.sh"
ln -sf "${DIR}/run.sh" "${BIN_DIR}/docflow-studio"

echo "==> Instalando entrada de escritorio…"
mkdir -p "$APPS_DIR"
cat > "${APPS_DIR}/docflow-studio.desktop" <<EOF
[Desktop Entry]
Type=Application
Name=DocFlow Studio
Comment=Conversión documental y gestión de PDF
Exec=${DIR}/run.sh
Terminal=false
Categories=Office;Utility;
EOF

echo
echo "Listo. Ejecuta 'docflow-studio' (o búscala en el menú de aplicaciones)."
echo "Sugerencia: instala también las dependencias de los backends:"
echo "  ${DIR}/backends/docflow/install.sh"
echo "  ${DIR}/backends/pdf-suite/install.sh"

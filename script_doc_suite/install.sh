#!/usr/bin/env bash
# install.sh — Instala docflow y sus dependencias en Arch Linux
# Bug #5 corregido: usa readlink -f para resolver symlinks antes de calcular SCRIPT_DIR
set -euo pipefail

BOLD='\033[1m'; GREEN='\033[0;32m'; YELLOW='\033[1;33m'
RED='\033[0;31m'; RESET='\033[0m'

# Resolver la ruta real del script aunque se llame desde un symlink
SCRIPT_DIR="$(cd "$(dirname "$(readlink -f "${BASH_SOURCE[0]}")")" && pwd)"

echo -e "${BOLD}Instalando docflow v2.0.0...${RESET}"
echo -e "${BOLD}Ruta del proyecto: ${SCRIPT_DIR}${RESET}\n"

# ---------------------------------------------------------------------------
# [1/4] Dependencias requeridas
# ---------------------------------------------------------------------------
echo -e "${BOLD}[1/4] Instalando dependencias requeridas (pandoc, python)...${RESET}"
if sudo pacman -S --needed --noconfirm pandoc python; then
    echo -e "${GREEN}  ✓ pandoc y python instalados.${RESET}"
else
    echo -e "${RED}  ✗ Error instalando dependencias requeridas.${RESET}"
    exit 1
fi

# ---------------------------------------------------------------------------
# [2/4] Dependencias opcionales
# ---------------------------------------------------------------------------
echo -e "\n${BOLD}[2/4] Dependencias opcionales...${RESET}"
echo -e "${YELLOW}  Se recomienda instalar libreoffice-still e imagemagick.${RESET}"
read -rp "  ¿Instalar libreoffice-still? [s/N] " resp_lo
if [[ "${resp_lo,,}" == "s" ]]; then
    sudo pacman -S --needed --noconfirm libreoffice-still
    echo -e "${GREEN}  ✓ LibreOffice instalado.${RESET}"
else
    echo -e "${YELLOW}  Omitido. Instala después: sudo pacman -S libreoffice-still${RESET}"
fi

read -rp "  ¿Instalar imagemagick? [s/N] " resp_im
if [[ "${resp_im,,}" == "s" ]]; then
    sudo pacman -S --needed --noconfirm imagemagick
    echo -e "${GREEN}  ✓ imagemagick instalado.${RESET}"
else
    echo -e "${YELLOW}  Omitido. Instala después: sudo pacman -S imagemagick${RESET}"
fi

read -rp "  ¿Instalar inotify-tools (recomendado para modo watch)? [s/N] " resp_ino
if [[ "${resp_ino,,}" == "s" ]]; then
    sudo pacman -S --needed --noconfirm inotify-tools
    echo -e "${GREEN}  ✓ inotify-tools instalado.${RESET}"
else
    echo -e "${YELLOW}  Omitido. El modo watch usará polling como fallback.${RESET}"
fi

# ---------------------------------------------------------------------------
# [3/4] Permisos de ejecución
# ---------------------------------------------------------------------------
echo -e "\n${BOLD}[3/4] Configurando permisos...${RESET}"
chmod +x "${SCRIPT_DIR}/main.sh"
chmod +x "${SCRIPT_DIR}/lib/"*.sh
echo -e "${GREEN}  ✓ Permisos configurados.${RESET}"

# ---------------------------------------------------------------------------
# [4/4] Crear enlace simbólico en /usr/local/bin
# ---------------------------------------------------------------------------
echo -e "\n${BOLD}[4/4] Instalando 'docflow' en /usr/local/bin...${RESET}"
if sudo ln -sf "${SCRIPT_DIR}/main.sh" /usr/local/bin/docflow; then
    echo -e "${GREEN}  ✓ Comando 'docflow' disponible globalmente.${RESET}"
else
    echo -e "${YELLOW}  No se pudo crear el enlace en /usr/local/bin.${RESET}"
    echo -e "${YELLOW}  Agrega manualmente al PATH en ~/.zshrc o ~/.bashrc:${RESET}"
    echo -e "    alias docflow='${SCRIPT_DIR}/main.sh'"
fi

# ---------------------------------------------------------------------------
# Resumen final
# ---------------------------------------------------------------------------
echo ""
echo -e "${GREEN}${BOLD}╔══════════════════════════════════════════╗${RESET}"
echo -e "${GREEN}${BOLD}║  ✓ Instalación completada correctamente  ║${RESET}"
echo -e "${GREEN}${BOLD}╚══════════════════════════════════════════╝${RESET}"
echo ""
echo -e "  Ejecuta ${BOLD}docflow --help${RESET} para ver todos los comandos."
echo -e "  Proyecto: ${SCRIPT_DIR}"
echo -e "  README:   ${SCRIPT_DIR}/README.md"
echo ""

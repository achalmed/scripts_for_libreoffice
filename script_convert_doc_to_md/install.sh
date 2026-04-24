#!/usr/bin/env bash
# install.sh — Instala doc2md y sus dependencias en Arch Linux
set -euo pipefail

BOLD='\033[1m'; GREEN='\033[0;32m'; YELLOW='\033[1;33m'
RED='\033[0;31m'; RESET='\033[0m'

echo -e "${BOLD}Instalando doc2md...${RESET}"

# Dependencias requeridas
echo -e "\n${BOLD}[1/3] Instalando dependencias requeridas (pandoc, python)...${RESET}"
sudo pacman -S --needed --noconfirm pandoc python

# Dependencias opcionales
echo -e "\n${BOLD}[2/3] Instalando dependencias opcionales...${RESET}"
echo -e "${YELLOW}Se recomienda instalar libreoffice e imagemagick para mejor compatibilidad.${RESET}"
read -rp "¿Instalar libreoffice-still e imagemagick? [s/N] " resp
if [[ "${resp,,}" == "s" ]]; then
    sudo pacman -S --needed --noconfirm libreoffice-still imagemagick
    echo -e "${GREEN}Instaladas.${RESET}"
else
    echo -e "${YELLOW}Omitido. Puedes instalarlas después:${RESET}"
    echo "  sudo pacman -S libreoffice-still imagemagick"
fi

# Instalar el script
echo -e "\n${BOLD}[3/3] Instalando doc2md en /usr/local/bin...${RESET}"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
sudo install -m 755 "${SCRIPT_DIR}/doc2md.sh" /usr/local/bin/doc2md

echo -e "\n${GREEN}${BOLD}¡Instalación completa!${RESET}"
echo -e "  Ejecuta ${BOLD}doc2md --help${RESET} para comenzar."
echo -e "  README: ${SCRIPT_DIR}/README.md"

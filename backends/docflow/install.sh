#!/usr/bin/env bash
# install.sh — Instalador de docflow para Arch, Debian/Ubuntu, Fedora y openSUSE.
# Instala dependencias (requeridas y opcionales a elección), enlaza el binario
# en el PATH y verifica la instalación con `docflow doctor`.
set -uo pipefail

DOCFLOW_ROOT="$(cd "$(dirname "$(readlink -f "${BASH_SOURCE[0]}")")" && pwd)"

BOLD=$'\033[1m'; GREEN=$'\033[0;32m'; YELLOW=$'\033[1;33m'; RED=$'\033[0;31m'; RESET=$'\033[0m'

say()  { printf '%b\n' "$*"; }
ok()   { say "${GREEN}✓${RESET} $*"; }
warn() { say "${YELLOW}!${RESET} $*"; }
die()  { say "${RED}✗${RESET} $*"; exit 1; }

# --- Detección de distro ----------------------------------------------------
detect_pkg() {
    if   command -v pacman  &>/dev/null; then echo pacman
    elif command -v apt-get &>/dev/null; then echo apt
    elif command -v dnf     &>/dev/null; then echo dnf
    elif command -v zypper  &>/dev/null; then echo zypper
    else echo ""; fi
}

pkg_install() {
    case "$PKG" in
        pacman) sudo pacman -S --needed --noconfirm "$@" ;;
        apt)    sudo apt-get install -y "$@" ;;
        dnf)    sudo dnf install -y "$@" ;;
        zypper) sudo zypper install -y "$@" ;;
        *)      warn "Gestor de paquetes desconocido; instala manualmente: $*"; return 1 ;;
    esac
}

# Paquetes por distro: requeridos y opcionales
pkgs_required() {
    case "$PKG" in
        pacman) echo "pandoc python libreoffice-still" ;;
        apt)    echo "pandoc python3 libreoffice" ;;
        dnf)    echo "pandoc python3 libreoffice" ;;
        zypper) echo "pandoc python3 libreoffice" ;;
    esac
}

pkgs_optional() {
    case "$PKG" in
        pacman) echo "qpdf ghostscript poppler imagemagick inotify-tools ocrmypdf texlive-xetex" ;;
        apt)    echo "qpdf ghostscript poppler-utils imagemagick inotify-tools ocrmypdf texlive-xetex" ;;
        dnf)    echo "qpdf ghostscript poppler-utils ImageMagick inotify-tools ocrmypdf texlive-xetex" ;;
        zypper) echo "qpdf ghostscript poppler-tools ImageMagick inotify-tools texlive-xetex" ;;
    esac
}

main() {
    say "${BOLD}docflow — instalador${RESET}"
    say ""

    PKG="$(detect_pkg)"
    [[ -n "$PKG" ]] && ok "Gestor de paquetes: ${PKG}" \
                    || warn "Gestor de paquetes no reconocido: se omite la instalación de dependencias."

    # --- Dependencias ---------------------------------------------------
    if [[ -n "$PKG" && "${1:-}" != "--no-deps" ]]; then
        say ""
        say "${BOLD}Dependencias requeridas:${RESET} $(pkgs_required)"
        read -r -p "¿Instalarlas ahora? [S/n] " reply
        if [[ ! "$reply" =~ ^[nN]$ ]]; then
            # shellcheck disable=SC2046
            pkg_install $(pkgs_required) && ok "Requeridas instaladas."
        fi

        say ""
        say "${BOLD}Dependencias opcionales:${RESET} $(pkgs_optional)"
        say "  (PDF avanzado, OCR, optimización de imágenes, modo watch)"
        read -r -p "¿Instalarlas también? [s/N] " reply
        if [[ "$reply" =~ ^[sS]$ ]]; then
            # shellcheck disable=SC2046
            pkg_install $(pkgs_optional) && ok "Opcionales instaladas."
        fi
    fi

    # --- Permisos y symlink ---------------------------------------------
    chmod +x "${DOCFLOW_ROOT}/bin/docflow" "${DOCFLOW_ROOT}"/lib/helpers/*.py

    local target_dir
    if [[ -w /usr/local/bin ]]; then
        target_dir="/usr/local/bin"
    elif [[ -d "${HOME}/.local/bin" ]]; then
        target_dir="${HOME}/.local/bin"
    else
        mkdir -p "${HOME}/.local/bin"
        target_dir="${HOME}/.local/bin"
    fi

    if [[ "$target_dir" == "/usr/local/bin" ]]; then
        ln -sf "${DOCFLOW_ROOT}/bin/docflow" "${target_dir}/docflow"
    else
        ln -sf "${DOCFLOW_ROOT}/bin/docflow" "${target_dir}/docflow"
        case ":$PATH:" in
            *":${target_dir}:"*) ;;
            *) warn "${target_dir} no está en tu PATH; añádelo a tu shell." ;;
        esac
    fi
    ok "Enlace creado: ${target_dir}/docflow → ${DOCFLOW_ROOT}/bin/docflow"

    # --- Completions ------------------------------------------------------
    local comp_dir="${XDG_DATA_HOME:-$HOME/.local/share}/bash-completion/completions"
    mkdir -p "$comp_dir"
    cp "${DOCFLOW_ROOT}/completions/docflow.bash" "${comp_dir}/docflow" 2>/dev/null \
        && ok "Autocompletado bash instalado."
    if [[ -d "${HOME}/.config/fish/completions" ]]; then
        cp "${DOCFLOW_ROOT}/completions/docflow.fish" "${HOME}/.config/fish/completions/" 2>/dev/null \
            && ok "Autocompletado fish instalado."
    fi

    # --- Configuración inicial -------------------------------------------
    local cfg="${XDG_CONFIG_HOME:-$HOME/.config}/docflow/config.toml"
    if [[ ! -f "$cfg" ]]; then
        "${DOCFLOW_ROOT}/bin/docflow" config init >/dev/null 2>&1 \
            && ok "Configuración creada: ${cfg}"
    fi

    # --- Verificación -----------------------------------------------------
    say ""
    say "${BOLD}Verificando instalación:${RESET}"
    "${DOCFLOW_ROOT}/bin/docflow" doctor || true

    say ""
    ok "${BOLD}docflow instalado.${RESET} Prueba: docflow to-md --help"
}

main "$@"

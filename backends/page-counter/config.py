"""
config.py — Configuración centralizada de pdf_page_counter.

Todas las rutas y constantes editables por el usuario viven aquí;
los módulos de lib/ nunca llevan rutas hardcodeadas.

Author : Edison Achalma (@achalmed)
Version: 2.0.0
"""

from pathlib import Path

SCRIPT_NAME = "pdf-page-counter"
SCRIPT_VERSION = "2.0.0"

# ---------------------------------------------------------------------------
# Rutas de los blogs
# ---------------------------------------------------------------------------
# Raíz del workspace. Los blogs "pub_<nombre>" son submódulos git del hub
# website-achalma y viven en RUTA_BASE_PUBLICACIONES / SUBDIR_PUBS
# (reorganización 2026-09-06; antes colgaban directamente de ~/Documents).
RUTA_BASE_PUBLICACIONES = Path.home() / "Documents"

# Subcarpeta (relativa a la ruta base) que contiene los pub_*
# Carpeta del hub (repo website-achalma) desde 2026-09-06
DIR_HUB = "04 index"
SUBDIR_PUBS = f"{DIR_HUB}/_pubs"

# Prefijo de las carpetas de blog dentro de SUBDIR_PUBS
PREFIJO_BLOG = "pub_"

# Blogs estándar (nombre lógico, sin prefijo; cada uno tiene su _site/)
BLOGS_ESTANDAR = [
    "actus-mercator",
    "aequilibria",
    "axiomata",
    "chaska",
    "dialectica-y-mercado",
    "epsilon-y-beta",
    "methodica",
    "numerus-scriptum",
    "optimums",
    "pecunia-fluxus",
    "res-publica",
]

# Blogs dentro de website-achalma (no tienen _site propio; cuelgan de
# website-achalma/_site). Nombre lógico -> ruta relativa a la ruta base.
BLOGS_WEBSITE_ACHALMA = {
    "blog": f"{DIR_HUB}/_site/blog",
    "teching": f"{DIR_HUB}/_site/teching",
}

# Alias que seleccionan todos los blogs de website-achalma a la vez
ALIAS_WEBSITE_ACHALMA = "website-achalma"

# ---------------------------------------------------------------------------
# Salida
# ---------------------------------------------------------------------------
# Directorio (relativo a este script) donde se guardan los reportes Excel
DIRECTORIO_EXCEL = "excel_databases"

# ---------------------------------------------------------------------------
# Códigos de salida estándar
# ---------------------------------------------------------------------------
EXIT_SUCCESS = 0
EXIT_ERROR = 1
EXIT_USAGE = 2
EXIT_NOT_FOUND = 3
EXIT_MISSING_DEPENDENCY = 5

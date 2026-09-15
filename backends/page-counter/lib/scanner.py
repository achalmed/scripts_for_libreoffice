"""
lib/scanner.py — Búsqueda de PDFs y conteo de páginas.

Usa pypdf (sucesor mantenido de PyPDF2) con fallback a PyPDF2 para
entornos antiguos. Los errores de lectura se registran con su causa
real en nivel DEBUG en vez de descartarse en silencio.
"""

import logging
import sys
from pathlib import Path
from typing import List, Tuple

from config import EXIT_MISSING_DEPENDENCY

# Import diferido: la ausencia de pypdf solo debe impedir escanear,
# no acciones que no leen PDFs (--listar, --help)
try:
    from pypdf import PdfReader
except ImportError:
    try:
        from PyPDF2 import PdfReader
    except ImportError:
        PdfReader = None


def require_pdf_library() -> None:
    """
    Aborts with exit code 5 if no PDF backend is available.

    Called right before scanning starts so dependency-free commands
    keep working on systems without pypdf.
    """
    if PdfReader is None:
        print("❌ Error: no está instalado pypdf (ni PyPDF2).")
        print("   Instala con: pip install pypdf")
        print("   o bien:      conda install -c conda-forge pypdf")
        sys.exit(EXIT_MISSING_DEPENDENCY)

# Estados posibles de cada archivo en el reporte
ESTADO_OK = "OK"
ESTADO_VACIO = "VACÍO"
ESTADO_ERROR = "ERROR"

# (ruta_relativa, numero_paginas, estado)
ResultadoPdf = Tuple[str, int, str]

logger = logging.getLogger("pdf-page-counter")


def count_pdf_pages(pdf_path: Path) -> int:
    """
    Counts the pages of a single PDF file.

    Args:
        pdf_path: Path to the PDF file.

    Returns:
        Page count (>= 0), or -1 when the file can't be read. The
        failure cause is logged at DEBUG level so --verbose reveals it —
        the v1.x swallowed the exception entirely.
    """
    try:
        return len(PdfReader(str(pdf_path)).pages)
    except Exception as read_error:  # pypdf lanza tipos heterogéneos según el corrupto
        logger.debug("No se pudo leer '%s': %s", pdf_path, read_error)
        return -1


def scan_directory_for_pdfs(directory: str, solo_index: bool = True) -> List[ResultadoPdf]:
    """
    Recursively finds PDFs under a directory and counts their pages.

    Args:
        directory: Root directory to search.
        solo_index: When True, only files named 'index.pdf' are matched.

    Returns:
        List of (relative_path, page_count, status) tuples. page_count
        is 0 for unreadable files (status ERROR) to keep totals additive.
    """
    require_pdf_library()
    results: List[ResultadoPdf] = []
    root = Path(directory)

    if not root.exists():
        print(f"⚠️  Directorio no encontrado: {directory}")
        return results

    pattern = "**/index.pdf" if solo_index else "**/*.pdf"
    pdf_files = sorted(root.glob(pattern))

    if not pdf_files:
        print(f"   ℹ️  No se encontraron archivos en: {root.name}")
        return results

    for pdf_path in pdf_files:
        if not pdf_path.is_file():
            continue
        relative_path = str(pdf_path.relative_to(root))
        page_count = count_pdf_pages(pdf_path)

        if page_count > 0:
            results.append((relative_path, page_count, ESTADO_OK))
            print(f"   ✓ {relative_path:<60} {page_count:>3} página(s)")
        elif page_count == 0:
            results.append((relative_path, 0, ESTADO_VACIO))
            print(f"   ⚠ {relative_path:<60} {'0':>3} página(s) [VACÍO]")
        else:
            results.append((relative_path, 0, ESTADO_ERROR))
            print(f"   ✗ {relative_path:<60} {'ERR':>3} [ERROR LECTURA]")

    return results

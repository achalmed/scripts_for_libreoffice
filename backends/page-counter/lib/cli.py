"""
lib/cli.py — Definición de la interfaz de línea de comandos.

Centralizar el parser aquí mantiene main.py limpio y hace la CLI
fácil de extender. Las flags son las mismas de la v1.x más
--verbose y --version.

Author : Edison Achalma (@achalmed)
Version: 2.0.0
"""

import argparse

from config import SCRIPT_VERSION


def build_argument_parser() -> argparse.ArgumentParser:
    """
    Defines all CLI arguments and flags for the tool.

    Returns:
        The configured ArgumentParser.
    """
    parser = argparse.ArgumentParser(
        description="Contador de páginas PDF para la familia de blogs Quarto",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos de uso:
  # Contar todos los blogs (solo index.pdf)
  %(prog)s

  # Contar blogs específicos
  %(prog)s -b actus-mercator aequilibria

  # Contar todos los PDFs (no solo index.pdf)
  %(prog)s --todos

  # Listar blogs disponibles
  %(prog)s --listar

  # Archivo de salida personalizado
  %(prog)s -o mi_reporte.xlsx
        """,
    )
    parser.add_argument(
        "-b", "--blogs",
        nargs="+",
        help="Blogs específicos a procesar (nombres separados por espacios)",
    )
    parser.add_argument(
        "-t", "--todos",
        action="store_true",
        help="Buscar todos los archivos PDF (no solo index.pdf)",
    )
    parser.add_argument(
        "-o", "--output",
        help="Nombre del archivo Excel de salida (se guarda en excel_databases/)",
    )
    parser.add_argument(
        "-l", "--listar",
        action="store_true",
        help="Listar todos los blogs disponibles y salir",
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Mostrar información detallada (incluye la causa de cada PDF ilegible)",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {SCRIPT_VERSION}",
    )
    return parser

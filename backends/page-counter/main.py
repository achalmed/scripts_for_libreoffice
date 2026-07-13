#!/usr/bin/env python3
"""
main.py — Contador de páginas PDF (punto de entrada).

Cuenta las páginas de los PDFs renderizados por la familia de blogs
Quarto y genera un reporte Excel con subtotales por blog.

Flujo:
  1. Parsear argumentos (lib/cli.py).
  2. Validar ruta base y nombres de blog (lib/validator.py).
  3. Escanear cada blog contando páginas (lib/scanner.py).
  4. Generar el reporte Excel (lib/excel_report.py).
  5. Mostrar el resumen final (lib/ui.py).

Este archivo solo orquesta; la lógica vive en lib/.

Author : Edison Achalma (@achalmed)
Version: 2.0.0
"""

import os
import sys
from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------------------------
# Bootstrap: garantiza que config y lib/ sean importables desde cualquier CWD
# ---------------------------------------------------------------------------
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import (
    DIRECTORIO_EXCEL,
    EXIT_ERROR,
    EXIT_NOT_FOUND,
    EXIT_SUCCESS,
    EXIT_USAGE,
    RUTA_BASE_PUBLICACIONES,
    SCRIPT_NAME,
)
from lib import ui
from lib.cli import build_argument_parser
from lib.excel_report import create_excel_report
from lib.logger import setup_logger
from lib.scanner import ESTADO_ERROR, ESTADO_OK, ESTADO_VACIO, scan_directory_for_pdfs
from lib.validator import find_unknown_blogs, resolve_blog_paths, validate_base_path


def ensure_excel_directory() -> Path:
    """Creates (if needed) and returns the Excel output directory."""
    excel_dir = Path(__file__).parent / DIRECTORIO_EXCEL
    excel_dir.mkdir(exist_ok=True)
    return excel_dir


def build_output_path(excel_dir: Path, custom_name: str, todos: bool) -> Path:
    """Resolves the report filename (custom or timestamped)."""
    if custom_name:
        return excel_dir / custom_name
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    scan_kind = "todos" if todos else "index"
    return excel_dir / f"conteo_paginas_{scan_kind}_{timestamp}.xlsx"


def scan_all_blogs(blog_paths: dict, solo_index: bool) -> tuple:
    """
    Runs the scanner over every blog and aggregates counters.

    Returns:
        (results_by_blog, total_files, total_pages, empty_count, error_count)
    """
    results_by_blog = {}
    total_files = total_pages = empty_count = error_count = 0

    for position, (blog_name, blog_path) in enumerate(blog_paths.items(), 1):
        print(f"\n📖 [{position}/{len(blog_paths)}] Procesando: {blog_name}")
        results = scan_directory_for_pdfs(blog_path, solo_index=solo_index)
        if not results:
            continue

        results_by_blog[blog_name] = results
        ok_files = sum(1 for _, _, status in results if status == ESTADO_OK)
        blog_pages = sum(pages for _, pages, status in results if status == ESTADO_OK)
        blog_empty = sum(1 for _, _, status in results if status == ESTADO_VACIO)
        blog_errors = sum(1 for _, _, status in results if status == ESTADO_ERROR)

        total_files += len(results)
        total_pages += blog_pages
        empty_count += blog_empty
        error_count += blog_errors
        print(f"   📊 Resumen: {ok_files} OK | {blog_errors} errores | {blog_pages} páginas")

    return results_by_blog, total_files, total_pages, empty_count, error_count


def main() -> int:
    """Orchestrates the full run; returns the process exit code."""
    args = build_argument_parser().parse_args()
    logger = setup_logger(SCRIPT_NAME, verbose=args.verbose)

    ui.print_header()

    if args.listar:
        ui.print_available_blogs()
        return EXIT_SUCCESS

    if not validate_base_path():
        logger.error(f"La ruta base no existe: {RUTA_BASE_PUBLICACIONES}")
        logger.error("Actualiza RUTA_BASE_PUBLICACIONES en config.py.")
        return EXIT_NOT_FOUND

    if args.blogs:
        unknown_blogs = find_unknown_blogs(args.blogs)
        if unknown_blogs:
            logger.error(f"Blogs no reconocidos: {', '.join(unknown_blogs)}")
            logger.error("Usa --listar para ver los nombres válidos.")
            return EXIT_USAGE

    blog_paths = resolve_blog_paths(args.blogs)
    if not blog_paths:
        logger.error("No se encontraron blogs para procesar (¿faltan renderizar los _site/?).")
        logger.error("Usa --listar para ver los blogs disponibles.")
        return EXIT_NOT_FOUND

    excel_dir = ensure_excel_directory()
    output_path = build_output_path(excel_dir, args.output, args.todos)

    ui.print_section("CONFIGURACIÓN")
    print(f"📁 Ruta base: {RUTA_BASE_PUBLICACIONES}")
    print(f"📊 Modo: {'Todos los PDFs' if args.todos else 'Solo index.pdf'}")
    print(f"📝 Blogs a procesar: {len(blog_paths)}")
    print(f"💾 Archivo de salida: {output_path.name}")

    ui.print_section("PROCESANDO BLOGS")
    results_by_blog, total_files, total_pages, empty_count, error_count = \
        scan_all_blogs(blog_paths, solo_index=not args.todos)

    if not results_by_blog:
        print("\n⚠️  No se encontraron archivos PDF en ningún blog.")
        return EXIT_NOT_FOUND

    ui.print_section("GENERANDO REPORTE")
    print("📝 Creando archivo Excel...")
    try:
        create_excel_report(results_by_blog, str(output_path),
                            solo_index=not args.todos)
    except (PermissionError, OSError) as save_error:
        logger.error(f"No se pudo guardar el reporte '{output_path}': {save_error}")
        return EXIT_ERROR
    print(f"✅ Archivo creado: {output_path}")

    ui.print_summary(total_files, total_pages, empty_count, error_count)
    print(f"💡 Tip: El archivo se guardó en: {excel_dir}/")
    print("💡 Tip: Usa --listar para ver todos los blogs disponibles\n")

    return EXIT_ERROR if error_count > 0 else EXIT_SUCCESS


if __name__ == "__main__":
    sys.exit(main())

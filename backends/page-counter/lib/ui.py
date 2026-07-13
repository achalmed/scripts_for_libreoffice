"""
lib/ui.py — Presentación en terminal.

Encabezados, secciones, listado de blogs y resumen final. Solo
formatea: no busca PDFs ni toca el filesystem más allá de comprobar
qué blogs existen para el listado.

Author : Edison Achalma (@achalmed)
Version: 2.0.0
"""

from datetime import datetime
from pathlib import Path

from config import (
    BLOGS_ESTANDAR,
    BLOGS_WEBSITE_ACHALMA,
    RUTA_BASE_PUBLICACIONES,
)
from lib.validator import blog_site_path


def print_header() -> None:
    """Prints the application banner."""
    print("\n" + "=" * 80)
    print("📊 PDF PAGE COUNTER - CONTADOR DE PÁGINAS PDF".center(80))
    print("=" * 80)
    print("👤 Autor: Edison Achalma")
    print("🏛️  Universidad Nacional de San Cristóbal de Huamanga")
    print("📅 Fecha:", datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
    print("=" * 80 + "\n")


def print_section(title: str) -> None:
    """Prints a section separator with its title."""
    print("\n" + "─" * 80)
    print(f"📌 {title}")
    print("─" * 80)


def print_summary(total_files: int, total_pages: int,
                  empty_count: int, error_count: int) -> None:
    """
    Prints the final run summary.

    Successful files are counted explicitly (OK only) — the v1.x
    derived them as total-errors, which silently counted empty PDFs
    as successes.
    """
    print("\n" + "=" * 80)
    print("📈 RESUMEN FINAL".center(80))
    print("=" * 80)
    print(f"✅ Archivos procesados exitosamente: {total_files - error_count - empty_count}")
    print(f"⚠️  Archivos vacíos (0 páginas): {empty_count}")
    print(f"❌ Archivos con errores: {error_count}")
    print(f"📄 Total de archivos analizados: {total_files}")
    print(f"📑 Total de páginas contadas: {total_pages:,}")
    print("=" * 80 + "\n")


def print_available_blogs() -> None:
    """Lists every configured blog and whether its _site exists."""
    print_section("BLOGS DISPONIBLES")

    print("\n📚 Blogs estándar:")
    for index, blog_name in enumerate(BLOGS_ESTANDAR, 1):
        status = "✓" if blog_site_path(blog_name).exists() else "✗"
        print(f"   {index:2d}. {status} {blog_name}")

    print("\n🌐 Blogs en website-achalma:")
    for section_name, relative_path in BLOGS_WEBSITE_ACHALMA.items():
        status = "✓" if (RUTA_BASE_PUBLICACIONES / relative_path).exists() else "✗"
        print(f"       {status} {section_name}")

    print("\n✓ = Disponible | ✗ = No encontrado (blog sin renderizar o ruta inexistente)\n")

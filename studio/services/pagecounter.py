"""Adaptador del backend page-counter (backends/page-counter, Python).

A diferencia de los backends Bash, este se importa directamente: es Python
puro y así el conteo corre dentro del proceso (PythonTask) sin subshell.
Se reutilizan sus módulos lib/ tal cual; aquí solo se orquesta, replicando
el flujo de su main.py pero reportando por `emit` en vez de print().
"""

from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path
from typing import Callable

from studio.core import paths


def _importar_backend(settings=None):
    """Inserta backends/page-counter en sys.path e importa sus módulos."""
    base = str(paths.pagecounter_dir(settings))
    if base not in sys.path:
        sys.path.insert(0, base)
    import config as pc_config                      # noqa: PLC0415
    from lib import scanner, validator             # noqa: PLC0415
    from lib.excel_report import create_excel_report  # noqa: PLC0415
    return pc_config, scanner, validator, create_excel_report


def blogs_disponibles(settings=None) -> list[str]:
    """Nombres lógicos de blog con _site/ renderizado ahora mismo.

    resolve_blog_paths usa claves de presentación «website-achalma/blog»
    para las secciones del sitio principal; aquí se normalizan a su nombre
    lógico seleccionable («blog», «teching»).
    """
    pc_config, _, validator, _ = _importar_backend(settings)
    prefijo = f"{pc_config.ALIAS_WEBSITE_ACHALMA}/"
    return sorted(clave.removeprefix(prefijo)
                  for clave in validator.resolve_blog_paths(None))


def blogs_conocidos(settings=None) -> list[str]:
    """Blogs seleccionables (sin el alias agregado website-achalma)."""
    pc_config, _, validator, _ = _importar_backend(settings)
    return sorted(n for n in validator.known_blog_names()
                  if n != pc_config.ALIAS_WEBSITE_ACHALMA)


def contar_paginas(blogs: list[str] | None, todos: bool, salida: str,
                   emit: Callable[[str], None], settings=None) -> int:
    """Cuenta páginas y genera el Excel. Devuelve código de salida.

    salida: ruta completa del .xlsx; si viene vacía se genera con timestamp
    en el excel_databases/ del backend (mismo criterio que su CLI).
    """
    pc_config, scanner, validator, create_excel_report = _importar_backend(settings)

    if not validator.validate_base_path():
        emit(f"La ruta base no existe: {pc_config.RUTA_BASE_PUBLICACIONES}")
        return pc_config.EXIT_NOT_FOUND

    if blogs:
        desconocidos = validator.find_unknown_blogs(blogs)
        if desconocidos:
            emit(f"Blogs no reconocidos: {', '.join(desconocidos)}")
            return pc_config.EXIT_USAGE

    rutas = validator.resolve_blog_paths(blogs or None)
    if not rutas:
        emit("No se encontraron blogs renderizados (¿faltan los _site/?).")
        return pc_config.EXIT_NOT_FOUND

    if salida:
        ruta_salida = Path(salida)
    else:
        excel_dir = paths.pagecounter_dir(settings) / pc_config.DIRECTORIO_EXCEL
        excel_dir.mkdir(exist_ok=True)
        marca = datetime.now().strftime("%Y%m%d_%H%M%S")
        tipo = "todos" if todos else "index"
        ruta_salida = excel_dir / f"conteo_paginas_{tipo}_{marca}.xlsx"

    resultados = {}
    total_archivos = total_paginas = vacios = errores = 0
    for i, (nombre, ruta) in enumerate(rutas.items(), 1):
        emit(f"[{i}/{len(rutas)}] Procesando: {nombre}")
        res = scanner.scan_directory_for_pdfs(ruta, solo_index=not todos)
        if not res:
            continue
        resultados[nombre] = res
        paginas = sum(p for _, p, s in res if s == scanner.ESTADO_OK)
        errs = sum(1 for _, _, s in res if s == scanner.ESTADO_ERROR)
        vacios += sum(1 for _, _, s in res if s == scanner.ESTADO_VACIO)
        total_archivos += len(res)
        total_paginas += paginas
        errores += errs
        emit(f"    {len(res)} archivos | {paginas} páginas | {errs} errores")

    if not resultados:
        emit("No se encontraron archivos PDF en ningún blog.")
        return pc_config.EXIT_NOT_FOUND

    try:
        create_excel_report(resultados, str(ruta_salida), solo_index=not todos)
    except (PermissionError, OSError) as exc:
        emit(f"No se pudo guardar el reporte '{ruta_salida}': {exc}")
        return pc_config.EXIT_ERROR

    emit(f"Reporte creado: {ruta_salida}")
    emit(f"Total: {total_archivos} archivos, {total_paginas} páginas, "
         f"{vacios} vacíos, {errores} errores")
    return pc_config.EXIT_ERROR if errores else pc_config.EXIT_SUCCESS

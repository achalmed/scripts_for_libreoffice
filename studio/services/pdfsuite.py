"""Adaptador del backend pdf-suite (backends/pdf-suite, CLI Bash).

Backend canónico de TODA la manipulación PDF de la aplicación: organizar
(merge/split/extract/rotate/reorder/delete), optimizar (compress/test/
optimize/repair/validate), seguridad (encrypt/decrypt/watermark), contenido
(convert/ocr) y metadatos. docflow también trae operaciones PDF, pero para
evitar duplicar la misma función en dos lugares solo se usa docflow para
PDF/A (que pdf-suite no implementa).
"""

from __future__ import annotations

from studio.core import paths


def _base(settings=None, dry_run: bool = False, verbose: bool = False,
          output: str = "", recursive: bool = False, force: bool = False) -> list[str]:
    cmd = paths.pdfsuite_cmd(settings)
    if verbose:
        cmd.append("--verbose")
    if dry_run:
        cmd.append("--dry-run")
    if force:
        cmd.append("--force")
    if recursive:
        cmd.append("--recursive")
    if output:
        cmd += ["--output", output]
    return cmd


def comprimir(archivos: list[str], metodo: str = "ebook", umbral: int = 0,
              recursive: bool = False, **kw) -> list[str]:
    cmd = _base(recursive=recursive, **kw) + ["compress", "-m", metodo]
    if umbral > 0:
        cmd += ["-t", str(umbral)]
    return cmd + archivos


def test_compresion(archivo: str, **kw) -> list[str]:
    return _base(**kw) + ["test", archivo]


def unir(archivos: list[str], **kw) -> list[str]:
    return _base(**kw) + ["merge"] + archivos


def dividir(archivo: str, paginas_por_parte: int = 0, **kw) -> list[str]:
    cmd = _base(**kw) + ["split"]
    if paginas_por_parte > 0:
        cmd += ["--pages", str(paginas_por_parte)]
    return cmd + [archivo]


def extraer(archivo: str, paginas: str, **kw) -> list[str]:
    """paginas admite rangos qpdf: '1-5', '1,3,7,10-12', '10-z'."""
    return _base(**kw) + ["extract", "--pages", paginas, archivo]


def rotar(archivo: str, angulo: int, paginas: str = "", **kw) -> list[str]:
    cmd = _base(**kw) + ["rotate", "--angle", str(angulo)]
    if paginas:
        cmd += ["--pages", paginas]
    return cmd + [archivo]


def reordenar(archivo: str, orden: str = "", **kw) -> list[str]:
    """Sin orden → invierte todas las páginas."""
    cmd = _base(**kw) + ["reorder"]
    if orden:
        cmd += ["--range", orden]
    return cmd + [archivo]


def eliminar_paginas(archivo: str, paginas: str, **kw) -> list[str]:
    return _base(**kw) + ["delete", "--pages", paginas, archivo]


def convertir(archivos: list[str], destino: str, dpi: int = 0, **kw) -> list[str]:
    """destino ∈ {png, jpg, svg, txt, html, pdf}. 'pdf' = imágenes → PDF."""
    cmd = _base(**kw) + ["convert", "--to", destino]
    if dpi > 0:
        cmd += ["--dpi", str(dpi)]
    return cmd + archivos


def extraer_imagenes(archivo: str, **kw) -> list[str]:
    return _base(**kw) + ["convert", "--extract-images", archivo]


def ocr(rutas: list[str], idioma: str = "spa", recursive: bool = False, **kw) -> list[str]:
    return _base(recursive=recursive, **kw) + ["ocr", "-l", idioma] + rutas


def ocr_escanear(directorio: str, **kw) -> list[str]:
    """Detecta qué PDFs necesitan OCR sin modificarlos."""
    return _base(**kw) + ["ocr", "--scan", directorio]


def cifrar(archivo: str, owner_pass: str, user_pass: str = "", **kw) -> list[str]:
    cmd = _base(**kw) + ["encrypt"]
    if user_pass:
        cmd += ["--user-pass", user_pass]
    return cmd + ["--owner-pass", owner_pass, archivo]


def descifrar(archivo: str, password: str, **kw) -> list[str]:
    return _base(**kw) + ["decrypt", "--password", password, archivo]


def marca_agua(archivo: str, texto: str, opacidad: float = 0.0,
               angulo: int = 0, color: str = "", **kw) -> list[str]:
    cmd = _base(**kw) + ["watermark", "--text", texto]
    if opacidad > 0:
        cmd += ["--opacity", str(opacidad)]
    if angulo:
        cmd += ["--angle", str(angulo)]
    if color:
        cmd += ["--color", color]
    return cmd + [archivo]


def numerar_paginas(archivo: str, **kw) -> list[str]:
    return _base(**kw) + ["protect", "--page-numbers", archivo]


def reparar(archivo: str, **kw) -> list[str]:
    return _base(**kw) + ["repair", archivo]


def validar(archivo: str, **kw) -> list[str]:
    return _base(**kw) + ["validate", archivo]


def optimizar(archivo: str, **kw) -> list[str]:
    return _base(**kw) + ["optimize", archivo]


def info(archivos: list[str], **kw) -> list[str]:
    return _base(**kw) + ["info"] + archivos


def metadatos_ver(archivo: str, json_out: bool = False, **kw) -> list[str]:
    cmd = _base(**kw) + ["metadata"]
    if json_out:
        cmd.append("--json")
    return cmd + [archivo]


def metadatos_escribir(archivo: str, titulo: str = "", autor: str = "",
                       asunto: str = "", keywords: str = "", **kw) -> list[str]:
    cmd = _base(**kw) + ["metadata"]
    if titulo:
        cmd += ["--set-title", titulo]
    if autor:
        cmd += ["--set-author", autor]
    if asunto:
        cmd += ["--set-subject", asunto]
    if keywords:
        cmd += ["--set-keywords", keywords]
    return cmd + [archivo]


def metadatos_desde_nombre(directorio: str, **kw) -> list[str]:
    return _base(recursive=True, **kw) + ["metadata", "--from-filename", directorio]


def metadatos_faltantes(directorio: str, **kw) -> list[str]:
    return _base(**kw) + ["metadata", "--check-missing", directorio]


def deps(**kw) -> list[str]:
    return _base(**kw) + ["deps"]

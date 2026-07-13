"""Adaptador del backend docflow (backends/docflow, CLI Bash).

Construye las líneas de comando exactas de docflow; NO reimplementa nada.
Dominio cubierto aquí: conversión documental (to-md, to-pdf, to-odf,
to-office), watch, reportes de sesión, doctor, caché, config y PDF/A
(única operación PDF que pdf-suite no ofrece).

El resto de operaciones PDF se canaliza por services/pdfsuite.py para que
cada función exista en un solo lugar de la aplicación.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from studio.core import paths


@dataclass
class OpcionesConversion:
    """Opciones comunes de un lote de conversión (mapean 1:1 a flags CLI)."""

    output_dir: str = ""
    formats: str = ""            # "docx,odt"
    exclude: list[str] = field(default_factory=list)
    jobs: int = 0                # 0 = defecto del backend (nproc)
    overwrite: bool = False
    dry_run: bool = False
    resume: bool = False
    no_cache: bool = False
    backup: bool = False
    delete: bool = False         # docflow pedirá confirmación salvo force
    force: bool = False
    verbose: bool = False
    report: str = ""             # "html,json"
    # Markdown
    flavor: str = ""
    no_metadata: bool = False
    no_notes: bool = False
    # Imágenes
    img_format: str = ""
    img_quality: int = 0
    img_max_width: int = 0
    img_optimize: bool = False
    # PDF
    pdf_engine: str = ""
    pdfa: bool = False
    compress: bool = False
    # OCR
    ocr: bool = False
    ocr_lang: str = ""

    def flags(self) -> list[str]:
        f: list[str] = []
        if self.output_dir:
            f += ["--output-dir", self.output_dir]
        if self.formats:
            f += ["--formats", self.formats]
        for ex in self.exclude:
            f += ["--exclude", ex]
        if self.jobs > 0:
            f += ["--jobs", str(self.jobs)]
        if self.overwrite:
            f.append("--overwrite")
        if self.dry_run:
            f.append("--dry-run")
        if self.resume:
            f.append("--resume")
        if self.no_cache:
            f.append("--no-cache")
        if self.backup:
            f.append("--backup")
        elif self.delete:
            f.append("--delete")
        if self.force:
            f.append("--force")
        if self.verbose:
            f.append("--verbose")
        if self.report:
            f += ["--report", self.report]
        if self.flavor:
            f += ["--flavor", self.flavor]
        if self.no_metadata:
            f.append("--no-metadata")
        if self.no_notes:
            f.append("--no-notes")
        if self.img_format:
            f += ["--img-format", self.img_format]
        if self.img_quality > 0:
            f += ["--img-quality", str(self.img_quality)]
        if self.img_max_width > 0:
            f += ["--img-max-width", str(self.img_max_width)]
        if self.img_optimize:
            f.append("--img-optimize")
        if self.pdf_engine:
            f += ["--pdf-engine", self.pdf_engine]
        if self.pdfa:
            f.append("--pdfa")
        if self.compress:
            f.append("--compress")
        if self.ocr:
            f.append("--ocr")
        if self.ocr_lang:
            f += ["--ocr-lang", self.ocr_lang]
        return f


def convertir(comando: str, rutas: list[str], opciones: OpcionesConversion, settings=None) -> list[str]:
    """comando ∈ {to-md, to-pdf, to-odf, to-office}."""
    return paths.docflow_cmd(settings) + [comando] + opciones.flags() + rutas


def pdf_a(archivos: list[str], settings=None) -> list[str]:
    """PDF/A-2b para archivado — exclusivo de docflow."""
    return paths.docflow_cmd(settings) + ["pdf", "pdfa"] + archivos


def watch(directorio: str, destino: str, formats: str = "", intervalo: int = 0, settings=None) -> list[str]:
    cmd = paths.docflow_cmd(settings) + ["watch", "--to", destino]
    if formats:
        cmd += ["--formats", formats]
    if intervalo > 0:
        cmd += ["--interval", str(intervalo)]
    return cmd + [directorio]


def report(formato: str, settings=None) -> list[str]:
    return paths.docflow_cmd(settings) + ["report", formato]


def doctor(settings=None) -> list[str]:
    return paths.docflow_cmd(settings) + ["doctor"]


def formats_table(settings=None) -> list[str]:
    return paths.docflow_cmd(settings) + ["formats"]


def cache(accion: str, settings=None) -> list[str]:
    """accion ∈ {stats, prune, clear}."""
    return paths.docflow_cmd(settings) + ["cache", accion]


def config(accion: str, settings=None) -> list[str]:
    """accion ∈ {init, show, path}."""
    return paths.docflow_cmd(settings) + ["config", accion]


def indice(rutas: list[str], settings=None) -> list[str]:
    return paths.docflow_cmd(settings) + ["index"] + rutas

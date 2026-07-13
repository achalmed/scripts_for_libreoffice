"""Dominio: Metadatos PDF (backend: pdf-suite / exiftool).

Ver metadatos XMP+DocInfo, editarlos, derivar el título del nombre de
archivo en lote y auditar carpetas con metadatos faltantes.
"""

from __future__ import annotations

from PySide6.QtWidgets import (
    QFormLayout, QGroupBox, QHBoxLayout, QLineEdit, QPushButton,
)

from studio.services import pdfsuite
from studio.ui.pages.base import BasePage, SelectorRuta

FILTRO_PDF = "PDF (*.pdf);;Todos los archivos (*)"


class MetadataPage(BasePage):
    titulo = "Metadatos"
    descripcion = ("Leer y escribir metadatos XMP y DocInfo de PDFs "
                   "(título, autor, asunto, palabras clave), individual o en lote.")

    def __init__(self, task_manager, settings, parent=None):
        super().__init__(task_manager, settings, parent)
        lay = self.layout_contenido()

        # --- un solo PDF -------------------------------------------------------
        caja_uno = QGroupBox("Un PDF")
        form = QFormLayout(caja_uno)
        self.archivo = SelectorRuta("metadata", settings, filtro=FILTRO_PDF)
        form.addRow("PDF:", self.archivo)

        self.titulo_pdf = QLineEdit()
        self.autor = QLineEdit()
        self.asunto = QLineEdit()
        self.keywords = QLineEdit()
        form.addRow("Título:", self.titulo_pdf)
        form.addRow("Autor:", self.autor)
        form.addRow("Asunto:", self.asunto)
        form.addRow("Palabras clave:", self.keywords)

        fila = QHBoxLayout()
        ver = QPushButton("Ver metadatos")
        ver.clicked.connect(self._ver)
        ver_json = QPushButton("Ver como JSON")
        ver_json.clicked.connect(lambda: self._ver(json_out=True))
        escribir = QPushButton("Escribir metadatos")
        escribir.clicked.connect(self._escribir)
        fila.addStretch(1)
        for b in (ver, ver_json, escribir):
            fila.addWidget(b)
        form.addRow(fila)
        lay.addWidget(caja_uno)

        # --- lote -----------------------------------------------------------------
        caja_lote = QGroupBox("Lote (carpeta)")
        form_lote = QFormLayout(caja_lote)
        self.carpeta = SelectorRuta("metadata-dir", settings, carpeta=True)
        form_lote.addRow("Carpeta:", self.carpeta)

        fila_lote = QHBoxLayout()
        desde_nombre = QPushButton("Título desde nombre de archivo (recursivo)")
        desde_nombre.clicked.connect(self._desde_nombre)
        faltantes = QPushButton("Auditar metadatos faltantes")
        faltantes.clicked.connect(self._faltantes)
        fila_lote.addStretch(1)
        fila_lote.addWidget(desde_nombre)
        fila_lote.addWidget(faltantes)
        form_lote.addRow(fila_lote)
        lay.addWidget(caja_lote)
        lay.addStretch(1)

    def recibir_rutas(self, rutas: list[str]) -> None:
        if rutas and rutas[0].lower().endswith(".pdf"):
            self.archivo.set_ruta(rutas[0])
        elif rutas:
            self.carpeta.set_ruta(rutas[0])

    def _ver(self, json_out: bool = False) -> None:
        if not self.archivo.ruta():
            self.aviso("Elige un PDF.")
            return
        self.ejecutar("Metadatos · ver",
                      pdfsuite.metadatos_ver(self.archivo.ruta(), json_out=json_out,
                                             settings=self.settings))

    def _escribir(self) -> None:
        if not self.archivo.ruta():
            self.aviso("Elige un PDF.")
            return
        if not any(c.text().strip() for c in (self.titulo_pdf, self.autor, self.asunto, self.keywords)):
            self.aviso("Completa al menos un campo de metadatos.")
            return
        self.ejecutar("Metadatos · escribir", pdfsuite.metadatos_escribir(
            self.archivo.ruta(),
            titulo=self.titulo_pdf.text().strip(),
            autor=self.autor.text().strip(),
            asunto=self.asunto.text().strip(),
            keywords=self.keywords.text().strip(),
            settings=self.settings))

    def _desde_nombre(self) -> None:
        if not self.carpeta.ruta():
            self.aviso("Elige una carpeta.")
            return
        self.ejecutar("Metadatos · título desde nombre",
                      pdfsuite.metadatos_desde_nombre(self.carpeta.ruta(), settings=self.settings))

    def _faltantes(self) -> None:
        if not self.carpeta.ruta():
            self.aviso("Elige una carpeta.")
            return
        self.ejecutar("Metadatos · auditar faltantes",
                      pdfsuite.metadatos_faltantes(self.carpeta.ruta(), settings=self.settings))

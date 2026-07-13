"""Dominio: Reportes.

- Reportes de sesión de docflow (html/csv/json/md de la última conversión).
- Conteo de páginas de los blogs Quarto (backend page-counter, en proceso)
  con salida Excel.
"""

from __future__ import annotations

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QFormLayout, QGroupBox, QHBoxLayout, QLabel,
    QLineEdit, QListWidget, QListWidgetItem, QPushButton,
)

from studio.services import docflow, pagecounter
from studio.ui.pages.base import BasePage


class ReportsPage(BasePage):
    titulo = "Reportes"
    descripcion = ("Reportes de la última sesión de conversión y conteo de páginas "
                   "de los PDFs renderizados por la familia de blogs Quarto.")

    def __init__(self, task_manager, settings, parent=None):
        super().__init__(task_manager, settings, parent)
        lay = self.layout_contenido()

        # --- reporte de sesión docflow -----------------------------------------
        caja_sesion = QGroupBox("Última sesión de conversión (docflow)")
        form = QFormLayout(caja_sesion)
        self.formato = QComboBox()
        self.formato.addItems(["html", "md", "csv", "json"])
        fila = QHBoxLayout()
        fila.addWidget(self.formato)
        generar = QPushButton("Generar reporte")
        generar.clicked.connect(self._reporte_sesion)
        fila.addWidget(generar)
        fila.addStretch(1)
        form.addRow("Formato:", fila)
        lay.addWidget(caja_sesion)

        # --- conteo de páginas Quarto ---------------------------------------------
        caja_conteo = QGroupBox("Conteo de páginas · blogs Quarto (Excel)")
        form_c = QFormLayout(caja_conteo)

        self.lista_blogs = QListWidget()
        self.lista_blogs.setMaximumHeight(160)
        form_c.addRow("Blogs:", self.lista_blogs)

        refrescar = QPushButton("Detectar blogs renderizados")
        refrescar.clicked.connect(self._cargar_blogs)
        form_c.addRow("", refrescar)

        self.todos_pdfs = QCheckBox("Todos los PDFs (no solo index.pdf)")
        form_c.addRow("", self.todos_pdfs)

        self.nombre_salida = QLineEdit()
        self.nombre_salida.setPlaceholderText("(vacío = nombre con timestamp)")
        form_c.addRow("Nombre del Excel:", self.nombre_salida)

        self._estado = QLabel("")
        fila_c = QHBoxLayout()
        contar = QPushButton("Contar páginas y generar Excel")
        contar.clicked.connect(self._contar)
        abrir = QPushButton("Abrir carpeta de reportes")
        abrir.clicked.connect(self._abrir_carpeta)
        fila_c.addStretch(1)
        fila_c.addWidget(contar)
        fila_c.addWidget(abrir)
        form_c.addRow(fila_c)
        form_c.addRow(self._estado)
        lay.addWidget(caja_conteo)
        lay.addStretch(1)

        self._blogs_cargados = False

    def showEvent(self, event) -> None:  # noqa: N802 — API Qt
        super().showEvent(event)
        if not self._blogs_cargados:
            self._cargar_blogs()

    # -- docflow ------------------------------------------------------------
    def _reporte_sesion(self) -> None:
        self.ejecutar(f"Reporte de sesión ({self.formato.currentText()})",
                      docflow.report(self.formato.currentText(), self.settings))

    # -- page counter -----------------------------------------------------------
    def _cargar_blogs(self) -> None:
        self.lista_blogs.clear()
        try:
            disponibles = set(pagecounter.blogs_disponibles(self.settings))
            conocidos = pagecounter.blogs_conocidos(self.settings)
        except Exception as exc:                     # noqa: BLE001 — backend ausente
            self._estado.setText(f"No se pudo cargar el backend page-counter: {exc}")
            return
        for nombre in conocidos:
            item = QListWidgetItem(
                nombre if nombre in disponibles else f"{nombre}  (sin renderizar)")
            item.setData(Qt.UserRole, nombre)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked if nombre in disponibles else Qt.Unchecked)
            self.lista_blogs.addItem(item)
        self._blogs_cargados = True
        self._estado.setText(f"{len(disponibles)} blog(s) con _site/ renderizado.")

    def _blogs_marcados(self) -> list[str]:
        marcados = []
        for i in range(self.lista_blogs.count()):
            item = self.lista_blogs.item(i)
            if item.checkState() == Qt.Checked:
                marcados.append(item.data(Qt.UserRole))
        return marcados

    def _contar(self) -> None:
        blogs = self._blogs_marcados()
        if not blogs:
            self.aviso("Marca al menos un blog.")
            return
        todos = self.todos_pdfs.isChecked()
        nombre = self.nombre_salida.text().strip()
        salida = ""
        if nombre:
            if not nombre.endswith(".xlsx"):
                nombre += ".xlsx"
            salida = str(self._carpeta_reportes() / nombre)
        settings = self.settings
        self.ejecutar_python(
            "Conteo de páginas Quarto",
            lambda emit: pagecounter.contar_paginas(blogs, todos, salida, emit, settings))

    def _carpeta_reportes(self):
        from studio.core import paths
        carpeta = paths.pagecounter_dir(self.settings) / "excel_databases"
        carpeta.mkdir(exist_ok=True)
        return carpeta

    def _abrir_carpeta(self) -> None:
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(self._carpeta_reportes())))

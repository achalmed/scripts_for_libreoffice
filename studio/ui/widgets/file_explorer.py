"""Explorador de archivos acoplable.

Doble clic (o «Usar como entrada» del menú contextual) envía la ruta a la
página activa mediante la señal `rutas_elegidas`; la ventana principal
decide a qué página entregarla.
"""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QDir, Qt, Signal
from PySide6.QtWidgets import (
    QFileSystemModel, QLineEdit, QMenu, QTreeView, QVBoxLayout, QWidget,
)


class FileExplorer(QWidget):
    rutas_elegidas = Signal(list)   # list[str]

    def __init__(self, raiz: str = "", parent=None):
        super().__init__(parent)
        self._modelo = QFileSystemModel(self)
        self._modelo.setRootPath(QDir.homePath())
        self._modelo.setFilter(QDir.AllDirs | QDir.Files | QDir.NoDotAndDotDot)

        self._ruta = QLineEdit()
        self._ruta.setPlaceholderText("Carpeta raíz del explorador…")
        self._ruta.returnPressed.connect(self._cambiar_raiz)

        self._arbol = QTreeView()
        self._arbol.setModel(self._modelo)
        self._arbol.setSelectionMode(QTreeView.ExtendedSelection)
        for col in (1, 2, 3):                       # solo nombre; tamaño/fecha estorban
            self._arbol.hideColumn(col)
        self._arbol.doubleClicked.connect(self._doble_clic)
        self._arbol.setContextMenuPolicy(Qt.CustomContextMenu)
        self._arbol.customContextMenuRequested.connect(self._menu_contextual)

        cuerpo = QVBoxLayout(self)
        cuerpo.setContentsMargins(4, 4, 4, 4)
        cuerpo.addWidget(self._ruta)
        cuerpo.addWidget(self._arbol)

        self.set_raiz(raiz or str(Path.home()))

    def set_raiz(self, ruta: str) -> None:
        if Path(ruta).is_dir():
            self._ruta.setText(ruta)
            self._arbol.setRootIndex(self._modelo.index(ruta))

    def raiz(self) -> str:
        return self._ruta.text()

    # -- internos ------------------------------------------------------------
    def _cambiar_raiz(self) -> None:
        self.set_raiz(self._ruta.text())

    def _seleccion(self) -> list[str]:
        indices = self._arbol.selectionModel().selectedRows(0)
        return [self._modelo.filePath(ix) for ix in indices]

    def _doble_clic(self, indice) -> None:
        ruta = self._modelo.filePath(indice)
        if Path(ruta).is_file():
            self.rutas_elegidas.emit([ruta])

    def _menu_contextual(self, punto) -> None:
        rutas = self._seleccion()
        if not rutas:
            return
        menu = QMenu(self)
        usar = menu.addAction(f"Usar como entrada ({len(rutas)} elemento/s)")
        accion = menu.exec(self._arbol.viewport().mapToGlobal(punto))
        if accion is usar:
            self.rutas_elegidas.emit(rutas)

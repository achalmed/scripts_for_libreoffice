"""Infraestructura común de las páginas de dominio.

- BasePage: acceso a TaskManager/QSettings, lanzamiento de tareas y
  recepción de rutas desde el explorador de archivos.
- ListaRutas: lista de archivos/carpetas de entrada con drag & drop.
- SelectorRuta: campo de ruta única con botón «Examinar».
"""

from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import (
    QFileDialog, QHBoxLayout, QLabel, QLineEdit, QListWidget, QMessageBox,
    QPushButton, QVBoxLayout, QWidget,
)

from studio.core.settings import last_dir, set_last_dir
from studio.core.tasks import ProcessTask, PythonTask


class BasePage(QWidget):
    """Página de dominio: título + contenido; sabe lanzar tareas."""

    titulo = "Página"
    descripcion = ""

    def __init__(self, task_manager, settings, parent=None):
        super().__init__(parent)
        self.tasks = task_manager
        self.settings = settings

        self._layout = QVBoxLayout(self)
        self._layout.setContentsMargins(16, 12, 16, 12)
        cabecera = QLabel(f"<h2>{self.titulo}</h2><p>{self.descripcion}</p>")
        cabecera.setWordWrap(True)
        self._layout.addWidget(cabecera)

    def layout_contenido(self) -> QVBoxLayout:
        return self._layout

    # -- lanzamiento de tareas ------------------------------------------------
    def ejecutar(self, titulo: str, comando: list[str]) -> ProcessTask:
        return self.tasks.submit(ProcessTask(titulo, comando))

    def ejecutar_python(self, titulo: str, fn) -> PythonTask:
        return self.tasks.submit(PythonTask(titulo, fn))

    def aviso(self, mensaje: str) -> None:
        QMessageBox.warning(self, self.titulo, mensaje)

    # -- integración con el explorador ---------------------------------------
    def recibir_rutas(self, rutas: list[str]) -> None:
        """La ventana principal entrega aquí las rutas elegidas en el
        explorador. Cada página decide qué hacer (defecto: nada)."""


class ListaRutas(QWidget):
    """Lista de rutas de entrada (archivos y/o carpetas) con drag & drop."""

    def __init__(self, contexto: str, settings, filtro: str = "", parent=None):
        super().__init__(parent)
        self._contexto = contexto
        self._settings = settings
        self._filtro = filtro or "Todos los archivos (*)"

        self.lista = QListWidget()
        self.lista.setSelectionMode(QListWidget.ExtendedSelection)
        self.setAcceptDrops(True)

        btn_archivos = QPushButton("Añadir archivos…")
        btn_carpeta = QPushButton("Añadir carpeta…")
        btn_quitar = QPushButton("Quitar")
        btn_limpiar = QPushButton("Limpiar")
        btn_archivos.clicked.connect(self._agregar_archivos)
        btn_carpeta.clicked.connect(self._agregar_carpeta)
        btn_quitar.clicked.connect(self._quitar_seleccion)
        btn_limpiar.clicked.connect(self.lista.clear)

        botones = QHBoxLayout()
        for b in (btn_archivos, btn_carpeta, btn_quitar, btn_limpiar):
            botones.addWidget(b)
        botones.addStretch(1)

        cuerpo = QVBoxLayout(self)
        cuerpo.setContentsMargins(0, 0, 0, 0)
        cuerpo.addWidget(self.lista)
        cuerpo.addLayout(botones)

    def rutas(self) -> list[str]:
        return [self.lista.item(i).text() for i in range(self.lista.count())]

    def agregar(self, rutas: list[str]) -> None:
        existentes = set(self.rutas())
        for r in rutas:
            if r not in existentes:
                self.lista.addItem(r)

    # -- internos --------------------------------------------------------------
    def _agregar_archivos(self) -> None:
        rutas, _ = QFileDialog.getOpenFileNames(
            self, "Elegir archivos", last_dir(self._settings, self._contexto), self._filtro)
        if rutas:
            set_last_dir(self._settings, self._contexto, str(Path(rutas[0]).parent))
            self.agregar(rutas)

    def _agregar_carpeta(self) -> None:
        ruta = QFileDialog.getExistingDirectory(
            self, "Elegir carpeta", last_dir(self._settings, self._contexto))
        if ruta:
            set_last_dir(self._settings, self._contexto, ruta)
            self.agregar([ruta])

    def _quitar_seleccion(self) -> None:
        for item in self.lista.selectedItems():
            self.lista.takeItem(self.lista.row(item))

    def dragEnterEvent(self, event) -> None:  # noqa: N802 — API Qt
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event) -> None:  # noqa: N802 — API Qt
        self.agregar([u.toLocalFile() for u in event.mimeData().urls() if u.isLocalFile()])


class SelectorRuta(QWidget):
    """LineEdit + «Examinar…» para una ruta única (archivo o carpeta)."""

    def __init__(self, contexto: str, settings, carpeta: bool = False,
                 filtro: str = "", guardar: bool = False, parent=None):
        super().__init__(parent)
        self._contexto = contexto
        self._settings = settings
        self._carpeta = carpeta
        self._guardar = guardar
        self._filtro = filtro or "Todos los archivos (*)"

        self.campo = QLineEdit()
        boton = QPushButton("Examinar…")
        boton.clicked.connect(self._examinar)

        fila = QHBoxLayout(self)
        fila.setContentsMargins(0, 0, 0, 0)
        fila.addWidget(self.campo, 1)
        fila.addWidget(boton)

    def ruta(self) -> str:
        return self.campo.text().strip()

    def set_ruta(self, ruta: str) -> None:
        self.campo.setText(ruta)

    def _examinar(self) -> None:
        inicio = last_dir(self._settings, self._contexto)
        if self._carpeta:
            ruta = QFileDialog.getExistingDirectory(self, "Elegir carpeta", inicio)
        elif self._guardar:
            ruta, _ = QFileDialog.getSaveFileName(self, "Guardar como", inicio, self._filtro)
        else:
            ruta, _ = QFileDialog.getOpenFileName(self, "Elegir archivo", inicio, self._filtro)
        if ruta:
            set_last_dir(self._settings, self._contexto,
                         ruta if self._carpeta else str(Path(ruta).parent))
            self.campo.setText(ruta)

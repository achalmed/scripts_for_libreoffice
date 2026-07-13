"""Ventana principal.

El shell (nav + stack + menús) se define en resources/ui/main_window.ui,
editable con Qt Designer (`pyside6-designer`). Este controlador lo carga
con QUiLoader y monta encima las páginas de dominio y los paneles
acoplables (explorador, consola, tareas, log).
"""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QFile, QSize, Qt
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import QDockWidget, QListWidgetItem, QMessageBox

from studio import APP_NAME, APP_VERSION
from studio.core import paths
from studio.ui.pages.conversion import ConversionPage
from studio.ui.pages.metadata import MetadataPage
from studio.ui.pages.pdf import PdfPage
from studio.ui.pages.reports import ReportsPage
from studio.ui.pages.system import SystemPage
from studio.ui.pages.watch import WatchPage
from studio.ui.widgets.console import ConsoleWidget
from studio.ui.widgets.file_explorer import FileExplorer
from studio.ui.widgets.log_viewer import LogViewer
from studio.ui.widgets.task_panel import TaskPanel

UI_FILE = Path(__file__).resolve().parents[1] / "resources" / "ui" / "main_window.ui"

PAGINAS = (ConversionPage, PdfPage, MetadataPage, ReportsPage, WatchPage, SystemPage)


class MainController:
    """Posee la ventana cargada del .ui y coordina páginas y docks."""

    def __init__(self, task_manager, settings):
        self.tasks = task_manager
        self.settings = settings

        archivo = QFile(str(UI_FILE))
        archivo.open(QFile.ReadOnly)
        self.win = QUiLoader().load(archivo)
        archivo.close()
        self.win.setWindowTitle(APP_NAME)

        self._montar_paginas()
        self._montar_docks()
        self._conectar_menu()
        self._restaurar_estado()
        self._avisar_backends_faltantes()

    # -- construcción ---------------------------------------------------------
    def _montar_paginas(self) -> None:
        self.paginas = []
        for clase in PAGINAS:
            pagina = clase(self.tasks, self.settings)
            self.paginas.append(pagina)
            self.win.pagesStack.addWidget(pagina)
            item = QListWidgetItem(pagina.titulo)
            item.setSizeHint(QSize(180, 40))
            self.win.navList.addItem(item)
        self.win.navList.currentRowChanged.connect(self._cambiar_pagina)
        self.win.navList.setCurrentRow(0)

    def _montar_docks(self) -> None:
        self.explorador = FileExplorer(str(self.settings.value("explorer/root", "")))
        self.explorador.rutas_elegidas.connect(self._entregar_rutas)
        self._dock(self.explorador, "Explorador de archivos", Qt.LeftDockWidgetArea)

        self.consola = ConsoleWidget(self.tasks)
        dock_consola = self._dock(self.consola, "Consola", Qt.BottomDockWidgetArea)

        self.panel_tareas = TaskPanel(self.tasks)
        dock_tareas = self._dock(self.panel_tareas, "Tareas", Qt.BottomDockWidgetArea)

        self.log = LogViewer()
        dock_log = self._dock(self.log, "Registro (log)", Qt.BottomDockWidgetArea)

        self.win.tabifyDockWidget(dock_consola, dock_tareas)
        self.win.tabifyDockWidget(dock_tareas, dock_log)
        dock_consola.raise_()

    def _dock(self, widget, titulo: str, area) -> QDockWidget:
        dock = QDockWidget(titulo, self.win)
        dock.setObjectName(titulo)
        dock.setWidget(widget)
        self.win.addDockWidget(area, dock)
        self.win.menuVer.addAction(dock.toggleViewAction())
        return dock

    def _conectar_menu(self) -> None:
        self.win.actionSalir.triggered.connect(self.win.close)
        self.win.actionAcercaDe.triggered.connect(self._acerca_de)

    # -- comportamiento -----------------------------------------------------------
    def _cambiar_pagina(self, indice: int) -> None:
        self.win.pagesStack.setCurrentIndex(indice)
        self.settings.setValue("ui/pagina", indice)
        self.win.statusbar.showMessage(self.paginas[indice].descripcion, 8000)

    def _entregar_rutas(self, rutas: list[str]) -> None:
        pagina = self.win.pagesStack.currentWidget()
        if pagina is not None:
            pagina.recibir_rutas(rutas)

    def _acerca_de(self) -> None:
        QMessageBox.about(
            self.win, APP_NAME,
            f"<b>{APP_NAME} v{APP_VERSION}</b><br>"
            "Frontend unificado de docflow, pdf-suite y page-counter.<br>"
            "Conversión documental · Gestión de PDF · Metadatos · Reportes.<br><br>"
            "Edison Achalma · <a href='https://github.com/achalmed'>github.com/achalmed</a>")

    def _avisar_backends_faltantes(self) -> None:
        faltan = paths.missing_backends(self.settings)
        if faltan:
            self.win.statusbar.showMessage(
                "Backends no encontrados: " + ", ".join(faltan) +
                " — revisa Sistema → Ajustes.", 0)

    # -- persistencia de estado de la ventana -------------------------------------
    def _restaurar_estado(self) -> None:
        geometria = self.settings.value("ui/geometria")
        estado = self.settings.value("ui/estado")
        if geometria is not None:
            self.win.restoreGeometry(geometria)
        if estado is not None:
            self.win.restoreState(estado)
        try:
            self.win.navList.setCurrentRow(int(self.settings.value("ui/pagina", 0)))
        except (TypeError, ValueError):
            pass

    def guardar_estado(self) -> None:
        self.settings.setValue("ui/geometria", self.win.saveGeometry())
        self.settings.setValue("ui/estado", self.win.saveState())
        self.settings.setValue("explorer/root", self.explorador.raiz())
        self.settings.sync()

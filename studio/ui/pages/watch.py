"""Dominio: Vigilancia de directorios (backend: docflow watch).

Convierte automáticamente lo que aparezca en una carpeta. La vigilancia es
una tarea de larga duración: se detiene desde aquí o desde el panel de
tareas (cancelar = SIGTERM al proceso watch).
"""

from __future__ import annotations

from PySide6.QtWidgets import (
    QComboBox, QFormLayout, QGroupBox, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QSpinBox,
)

from studio.services import docflow
from studio.ui.pages.base import BasePage, SelectorRuta

DESTINOS = ("pdf", "md", "odf", "office")


class WatchPage(BasePage):
    titulo = "Vigilancia"
    descripcion = ("Vigila una carpeta y convierte al vuelo los documentos que "
                   "aparezcan (inotify si está disponible; si no, sondeo periódico).")

    def __init__(self, task_manager, settings, parent=None):
        super().__init__(task_manager, settings, parent)
        lay = self.layout_contenido()

        caja = QGroupBox("Configuración")
        form = QFormLayout(caja)

        self.directorio = SelectorRuta("watch", settings, carpeta=True)
        form.addRow("Carpeta vigilada:", self.directorio)

        self.destino = QComboBox()
        self.destino.addItems(DESTINOS)
        form.addRow("Convertir a:", self.destino)

        self.formatos = QLineEdit()
        self.formatos.setPlaceholderText("ej.: docx,pptx (vacío = todos)")
        form.addRow("Solo extensiones:", self.formatos)

        self.intervalo = QSpinBox()
        self.intervalo.setRange(0, 3600)
        self.intervalo.setSuffix(" s")
        self.intervalo.setSpecialValueText("defecto")
        form.addRow("Intervalo de sondeo:", self.intervalo)
        lay.addWidget(caja)

        self._estado = QLabel("Sin vigilancia activa.")
        fila = QHBoxLayout()
        fila.addWidget(self._estado, 1)
        self.btn_iniciar = QPushButton("Iniciar vigilancia")
        self.btn_iniciar.clicked.connect(self._iniciar)
        self.btn_detener = QPushButton("Detener")
        self.btn_detener.setEnabled(False)
        self.btn_detener.clicked.connect(self._detener)
        fila.addWidget(self.btn_iniciar)
        fila.addWidget(self.btn_detener)
        lay.addLayout(fila)
        lay.addStretch(1)

        self._tarea = None

    def recibir_rutas(self, rutas: list[str]) -> None:
        if rutas:
            self.directorio.set_ruta(rutas[0])

    def _iniciar(self) -> None:
        if not self.directorio.ruta():
            self.aviso("Elige la carpeta a vigilar.")
            return
        comando = docflow.watch(
            self.directorio.ruta(),
            self.destino.currentText(),
            formats=self.formatos.text().strip(),
            intervalo=self.intervalo.value(),
            settings=self.settings,
        )
        self._tarea = self.ejecutar(f"Vigilancia de {self.directorio.ruta()}", comando)
        self._tarea.finished.connect(self._al_terminar)
        self._estado.setText(f"Vigilando {self.directorio.ruta()} → {self.destino.currentText()}")
        self.btn_iniciar.setEnabled(False)
        self.btn_detener.setEnabled(True)

    def _detener(self) -> None:
        if self._tarea is not None:
            self._tarea.cancel()

    def _al_terminar(self, _code: int) -> None:
        self._estado.setText("Sin vigilancia activa.")
        self.btn_iniciar.setEnabled(True)
        self.btn_detener.setEnabled(False)
        self._tarea = None

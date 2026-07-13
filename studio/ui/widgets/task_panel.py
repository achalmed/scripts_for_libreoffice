"""Panel de tareas: estado en vivo de cada operación lanzada."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QHBoxLayout, QHeaderView, QPushButton, QTableWidget, QTableWidgetItem,
    QVBoxLayout, QWidget,
)

COLUMNAS = ("Tarea", "Estado", "Duración", "Código")


class TaskPanel(QWidget):
    def __init__(self, task_manager, parent=None):
        super().__init__(parent)
        self._manager = task_manager
        self._filas: dict[object, int] = {}

        self._tabla = QTableWidget(0, len(COLUMNAS))
        self._tabla.setHorizontalHeaderLabels(COLUMNAS)
        self._tabla.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self._tabla.setEditTriggers(QTableWidget.NoEditTriggers)
        self._tabla.setSelectionBehavior(QTableWidget.SelectRows)

        cancelar = QPushButton("Cancelar seleccionada")
        cancelar.clicked.connect(self._cancelar_seleccion)

        botones = QHBoxLayout()
        botones.addStretch(1)
        botones.addWidget(cancelar)

        cuerpo = QVBoxLayout(self)
        cuerpo.setContentsMargins(4, 4, 4, 4)
        cuerpo.addWidget(self._tabla)
        cuerpo.addLayout(botones)

        task_manager.task_added.connect(self._agregar)

    def _agregar(self, task) -> None:
        fila = self._tabla.rowCount()
        self._tabla.insertRow(fila)
        self._filas[task] = fila
        self._set(fila, 0, task.titulo)
        self._set(fila, 1, task.estado)
        self._set(fila, 2, "—")
        self._set(fila, 3, "—")
        task.started.connect(lambda t=task: self._refrescar(t))
        task.finished.connect(lambda _c, t=task: self._refrescar(t))

    def _refrescar(self, task) -> None:
        fila = self._filas[task]
        self._set(fila, 1, task.estado)
        if task.duracion is not None:
            self._set(fila, 2, f"{task.duracion:.1f}s")
        if task.exit_code is not None:
            self._set(fila, 3, str(task.exit_code))

    def _set(self, fila: int, col: int, texto: str) -> None:
        item = QTableWidgetItem(texto)
        item.setTextAlignment(Qt.AlignCenter if col else Qt.AlignVCenter | Qt.AlignLeft)
        self._tabla.setItem(fila, col, item)

    def _cancelar_seleccion(self) -> None:
        filas = {ix.row() for ix in self._tabla.selectionModel().selectedRows()}
        for task, fila in self._filas.items():
            if fila in filas and task.estado == "ejecutando" and task.cancelable():
                task.cancel()

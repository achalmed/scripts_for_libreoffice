"""Consola integrada: salida en vivo de todas las tareas.

Se suscribe al TaskManager, así cualquier tarea (de cualquier página)
aparece aquí sin que las páginas tengan que conocer la consola.
"""

from __future__ import annotations

from PySide6.QtGui import QFont, QTextCursor
from PySide6.QtWidgets import QHBoxLayout, QPlainTextEdit, QPushButton, QVBoxLayout, QWidget

MAX_BLOQUES = 8000  # líneas retenidas; evita crecer sin límite en lotes enormes


class ConsoleWidget(QWidget):
    def __init__(self, task_manager, parent=None):
        super().__init__(parent)
        self._texto = QPlainTextEdit(readOnly=True)
        self._texto.setMaximumBlockCount(MAX_BLOQUES)
        fuente = QFont("Monospace")
        fuente.setStyleHint(QFont.TypeWriter)
        self._texto.setFont(fuente)

        limpiar = QPushButton("Limpiar")
        limpiar.clicked.connect(self._texto.clear)

        botones = QHBoxLayout()
        botones.addStretch(1)
        botones.addWidget(limpiar)

        cuerpo = QVBoxLayout(self)
        cuerpo.setContentsMargins(4, 4, 4, 4)
        cuerpo.addWidget(self._texto)
        cuerpo.addLayout(botones)

        task_manager.task_added.connect(self._attach)

    def _attach(self, task) -> None:
        self.append_line(f"— {task.titulo} —")
        task.output.connect(self.append_line)
        task.finished.connect(
            lambda code, t=task: self.append_line(f"— {t.titulo}: {t.estado} (código {code}) —")
        )

    def append_line(self, linea: str) -> None:
        self._texto.appendPlainText(linea)
        self._texto.moveCursor(QTextCursor.End)

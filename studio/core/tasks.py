"""Gestor de tareas en segundo plano.

Toda operación de backend corre como una Task para que la interfaz nunca
se bloquee:

- ProcessTask → CLIs Bash (docflow, pdf-suite) vía QProcess, con salida
  en streaming hacia la consola integrada y cancelación real (SIGTERM).
- PythonTask  → backends importados en Python (page-counter) en un QThread.

TaskManager centraliza el registro: el panel de tareas, la consola y el
log de la app se suscriben a sus señales; las páginas solo crean tareas.
"""

from __future__ import annotations

import logging
import time
from collections.abc import Callable

from PySide6.QtCore import QObject, QProcess, QThread, Signal

log = logging.getLogger("studio.tasks")

ESTADOS = ("pendiente", "ejecutando", "completada", "error", "cancelada")


class Task(QObject):
    """Unidad de trabajo observable. No se instancia directamente."""

    output = Signal(str)          # línea de salida (stdout+stderr mezclados)
    started = Signal()
    finished = Signal(int)        # código de salida (-1 = cancelada/fallo interno)

    def __init__(self, titulo: str, parent=None):
        super().__init__(parent)
        self.titulo = titulo
        self.estado = "pendiente"
        self.inicio: float | None = None
        self.duracion: float | None = None
        self.exit_code: int | None = None

    def start(self) -> None:
        raise NotImplementedError

    def cancel(self) -> None:
        raise NotImplementedError

    def cancelable(self) -> bool:
        return True

    # -- helpers comunes ----------------------------------------------------
    def _mark_started(self) -> None:
        self.estado = "ejecutando"
        self.inicio = time.monotonic()
        self.started.emit()

    def _mark_finished(self, code: int, cancelada: bool = False) -> None:
        if self.inicio is not None:
            self.duracion = time.monotonic() - self.inicio
        self.exit_code = code
        if cancelada:
            self.estado = "cancelada"
        else:
            self.estado = "completada" if code == 0 else "error"
        self.finished.emit(code)


class ProcessTask(Task):
    """Ejecuta un comando externo con QProcess (no bloquea el hilo de UI)."""

    def __init__(self, titulo: str, comando: list[str], cwd: str | None = None, parent=None):
        super().__init__(titulo, parent)
        self.comando = comando
        self._cancelada = False
        self._proc = QProcess(self)
        self._proc.setProcessChannelMode(QProcess.MergedChannels)
        if cwd:
            self._proc.setWorkingDirectory(cwd)
        self._proc.readyReadStandardOutput.connect(self._on_output)
        self._proc.finished.connect(self._on_finished)
        self._proc.errorOccurred.connect(self._on_error)
        self._buffer = ""

    def start(self) -> None:
        self._mark_started()
        self.output.emit("$ " + " ".join(self.comando))
        self._proc.start(self.comando[0], self.comando[1:])

    def cancel(self) -> None:
        if self._proc.state() != QProcess.NotRunning:
            self._cancelada = True
            self._proc.terminate()          # SIGTERM: deja limpiar traps de Bash
            if not self._proc.waitForFinished(3000):
                self._proc.kill()

    def _on_output(self) -> None:
        datos = self._proc.readAllStandardOutput().data().decode("utf-8", "replace")
        self._buffer += datos
        while "\n" in self._buffer:
            linea, self._buffer = self._buffer.split("\n", 1)
            self.output.emit(linea)

    def _on_finished(self, code: int, _status) -> None:
        if self._buffer:
            self.output.emit(self._buffer)
            self._buffer = ""
        self._mark_finished(code, cancelada=self._cancelada)

    def _on_error(self, error) -> None:
        if error == QProcess.FailedToStart:
            self.output.emit(f"[studio] No se pudo iniciar: {self.comando[0]}")
            self._mark_finished(-1)


class _PythonWorker(QObject):
    """Corre el callable en el hilo del QThread y reporta por señales."""

    line = Signal(str)
    done = Signal(int)

    def __init__(self, fn: Callable[[Callable[[str], None]], int]):
        super().__init__()
        self._fn = fn

    def run(self) -> None:
        try:
            code = self._fn(self.line.emit)
        except Exception as exc:                      # noqa: BLE001 — se reporta al usuario
            self.line.emit(f"[studio] Error inesperado: {exc}")
            code = 1
        self.done.emit(int(code or 0))


class PythonTask(Task):
    """Ejecuta un callable Python en un QThread.

    El callable recibe una función `emit(línea)` para reportar progreso y
    devuelve un código de salida entero. No es cancelable (los backends
    Python actuales no tienen puntos de cancelación cooperativa).
    """

    def __init__(self, titulo: str, fn: Callable[[Callable[[str], None]], int], parent=None):
        super().__init__(titulo, parent)
        self._thread = QThread(self)
        self._worker = _PythonWorker(fn)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.line.connect(self.output)
        self._worker.done.connect(self._on_done)

    def start(self) -> None:
        self._mark_started()
        self._thread.start()

    def cancel(self) -> None:
        pass

    def cancelable(self) -> bool:
        return False

    def _on_done(self, code: int) -> None:
        self._thread.quit()
        self._thread.wait(2000)
        self._mark_finished(code)


class TaskManager(QObject):
    """Registro central de tareas de la sesión."""

    task_added = Signal(object)     # Task

    def __init__(self, parent=None):
        super().__init__(parent)
        self.tasks: list[Task] = []

    def submit(self, task: Task) -> Task:
        self.tasks.append(task)
        task.output.connect(lambda linea, t=task: log.info("[%s] %s", t.titulo, linea))
        task.finished.connect(
            lambda code, t=task: log.info("[%s] terminó (estado=%s, código=%s)", t.titulo, t.estado, code)
        )
        self.task_added.emit(task)
        task.start()
        return task

    def running(self) -> list[Task]:
        return [t for t in self.tasks if t.estado == "ejecutando"]

    def cancel_all(self) -> None:
        for t in self.running():
            t.cancel()

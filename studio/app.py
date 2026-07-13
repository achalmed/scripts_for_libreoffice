"""Arranque de la aplicación: logging, QApplication y ventana principal."""

from __future__ import annotations

import logging
import logging.handlers
import signal
import sys

from PySide6.QtWidgets import QApplication

from studio import APP_ID, APP_NAME, APP_ORG
from studio.core import paths
from studio.core.settings import app_settings
from studio.core.tasks import TaskManager
from studio.ui.main_window import MainController


def _configurar_logging() -> None:
    paths.ensure_state_dir()
    handler = logging.handlers.RotatingFileHandler(
        paths.LOG_FILE, maxBytes=2_000_000, backupCount=3, encoding="utf-8")
    handler.setFormatter(logging.Formatter(
        "%(asctime)s %(levelname)s %(name)s: %(message)s", "%Y-%m-%d %H:%M:%S"))
    raiz = logging.getLogger()
    raiz.setLevel(logging.INFO)
    raiz.addHandler(handler)


def main() -> int:
    _configurar_logging()
    logging.getLogger("studio").info("Iniciando %s", APP_NAME)

    app = QApplication(sys.argv)
    app.setApplicationName(APP_ID)
    app.setApplicationDisplayName(APP_NAME)
    app.setOrganizationName(APP_ORG)
    # Ctrl-C en la terminal cierra la app en lugar de quedar ignorado
    signal.signal(signal.SIGINT, signal.SIG_DFL)

    settings = app_settings()
    tasks = TaskManager()
    controlador = MainController(tasks, settings)

    def al_salir() -> None:
        tasks.cancel_all()
        controlador.guardar_estado()

    app.aboutToQuit.connect(al_salir)
    controlador.win.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())

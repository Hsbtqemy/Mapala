"""Point d'entrée de l'application Mapala."""

from __future__ import annotations

import logging
import sys
import threading
from pathlib import Path

import pandas  # noqa: F401
from PySide6.QtWidgets import QApplication

from mapala.main_window import MainWindow


_LOG_PATH = Path.home() / ".mapala_error.log"


def _init_logging() -> logging.Logger:
    logger = logging.getLogger("mapala")
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)
    handler = logging.FileHandler(_LOG_PATH, encoding="utf-8")
    formatter = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    logger.info("Mapala started")
    return logger


def _install_exception_hooks() -> None:
    def _handle_exception(exc_type, exc, tb) -> None:
        logger = _init_logging()
        logger.error("Unhandled exception", exc_info=(exc_type, exc, tb))
        sys.__excepthook__(exc_type, exc, tb)

    def _thread_exception(args: threading.ExceptHookArgs) -> None:
        _handle_exception(args.exc_type, args.exc_value, args.exc_traceback)

    sys.excepthook = _handle_exception
    threading.excepthook = _thread_exception


def main() -> int:
    _init_logging()
    _install_exception_hooks()
    app = QApplication(sys.argv)
    app.setApplicationName("Mapala")
    app.setOrganizationName("Mapala")
    win = MainWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())

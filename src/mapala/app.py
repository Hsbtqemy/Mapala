"""Point d'entrée de l'application Mapala."""

from __future__ import annotations

import sys

import pandas  # noqa: F401
from PySide6.QtWidgets import QApplication

from mapala.main_window import MainWindow


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("Mapala")
    app.setOrganizationName("Mapala")
    win = MainWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())

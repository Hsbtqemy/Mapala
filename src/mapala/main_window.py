"""Fenêtre principale Mapala."""

from __future__ import annotations

from PySide6.QtGui import QCloseEvent
from PySide6.QtWidgets import QMainWindow, QMessageBox

from mapala.screens.mapping_screen import MappingScreen


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Mapala")
        self.setMinimumSize(1000, 700)
        self.resize(1200, 800)
        self._screen = MappingScreen()
        self.setCentralWidget(self._screen)
        self._maybe_restore_session()

    def _maybe_restore_session(self) -> None:
        if not self._screen.has_autosave():
            return
        box = QMessageBox(self)
        box.setWindowTitle("Reprendre la session ?")
        box.setText("Reprendre la session ?")
        resume_btn = box.addButton("Reprendre", QMessageBox.ButtonRole.AcceptRole)
        box.addButton("Ignorer", QMessageBox.ButtonRole.RejectRole)
        box.exec()
        if box.clickedButton() == resume_btn:
            self._screen.restore_autosave()

    def closeEvent(self, event: QCloseEvent) -> None:  # type: ignore[override]
        self._screen.save_autosave()
        super().closeEvent(event)

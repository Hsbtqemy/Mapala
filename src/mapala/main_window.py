"""Fenêtre principale Mapala."""

from __future__ import annotations

from PySide6.QtWidgets import QMainWindow

from mapala.screens.mapping_screen import MappingScreen


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Mapala")
        self.setMinimumSize(1000, 700)
        self.resize(1200, 800)
        self._screen = MappingScreen()
        self.setCentralWidget(self._screen)

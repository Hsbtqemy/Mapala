"""Écran unique de mapping Template ↔ Source."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd
from PySide6.QtCore import QEvent, QObject, QItemSelectionModel, QThread, Qt, Signal
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMenu,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QSpinBox,
    QSplitter,
    QStyle,
    QTableWidget,
    QTableWidgetItem,
    QToolButton,
    QVBoxLayout,
    QWidget,
    QApplication,
)

from mapala.io_excel import (
    SUPPORTED_INPUT_FILTER,
    SUPPORTED_OUTPUT_FILTER,
    ExcelFileError,
    list_sheets,
    load_sheet,
    load_sheet_raw,
    save_output,
)
from mapala.template_builder import TemplateBuilderConfig, ZoneSpec, _build_zone_output, _format_value, build_output


CONCAT_MENU_VALUE = "__CONCAT__"
SOURCE_ORDER_ORIGIN = "origin"
SOURCE_ORDER_AZ = "az"
SOURCE_ORDER_ZA = "za"
SOURCE_ORDER_VALUE = "value"
SOURCE_ORDER_USAGE = "usage"
SOURCE_ORDER_MANUAL = "manual"


class _ExportWorker(QObject):
    finished = Signal(bool, str)

    def __init__(
        self,
        config: TemplateBuilderConfig,
        output_path: str,
        *,
        csv_separator: str = ";",
        drop_empty_columns: bool = False,
    ) -> None:
        super().__init__()
        self._config = config
        self._output_path = output_path
        self._csv_separator = csv_separator
        self._drop_empty_columns = drop_empty_columns

    def run(self) -> None:
        try:
            output = build_output(self._config)
            save_output(
                self._output_path,
                output,
                header=False,
                index=False,
                csv_separator=self._csv_separator,
                drop_empty_columns=self._drop_empty_columns,
            )
        except Exception as e:
            self.finished.emit(False, str(e))
            return
        self.finished.emit(True, "")


def _parse_int_list(text: str) -> list[int]:
    items: list[int] = []
    for part in text.replace(";", ",").split(","):
        part = part.strip()
        if not part:
            continue
        try:
            items.append(int(part))
        except ValueError:
            continue
    return items


class _ConcatSourceWidget(QWidget):
    changed = Signal()

    def __init__(self, source_cols: list[str], col: str = "", prefix: str = "", parent: QWidget | None = None) -> None:
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)
        self._col_combo = QComboBox()
        self._prefix_edit = QLineEdit()
        self._prefix_edit.setPlaceholderText("Préfixe (optionnel)")
        layout.addWidget(self._col_combo, 2)
        layout.addWidget(self._prefix_edit, 3)
        self._col_combo.currentTextChanged.connect(lambda _text=None: self.changed.emit())
        self._prefix_edit.editingFinished.connect(lambda: self.changed.emit())
        self.refresh_source_cols(source_cols, current=col)
        if prefix:
            self._prefix_edit.setText(prefix)

    def refresh_source_cols(self, source_cols: list[str], current: str | None = None) -> None:
        cur = current if current is not None else self._col_combo.currentText()
        items = list(source_cols)
        if cur and cur not in items:
            items.insert(0, cur)
        self._col_combo.clear()
        self._col_combo.addItems(items)
        if cur:
            self._col_combo.setCurrentText(cur)

    def get_data(self) -> tuple[str, str]:
        return self._col_combo.currentText().strip(), self._prefix_edit.text()


class MappingScreen(QWidget):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._template_path: str = ""
        self._source_path: str = ""
        self._template_df_raw: pd.DataFrame | None = None
        self._source_df: pd.DataFrame | None = None
        self._zone: dict[str, Any] = {
            "name": "Zone",
            "row_start": 1,
            "row_end": None,
            "col_start": 1,
            "col_end": None,
            "header": {"tech_row": 1, "label_rows": []},
            "field_mappings": [],
        }
        self._current_mapping_cols: list[int] = []
        self._current_mapping_labels: list[str] = []
        self._current_preview_col: int | None = None
        self._mapping_concat_loading = False
        self._preview_combo_by_col: dict[int, QComboBox] = {}
        self._preview_concat_btn_by_col: dict[int, QToolButton] = {}
        self._source_order_mode = SOURCE_ORDER_ORIGIN
        self._source_default_order: list[str] = []
        self._source_manual_order: list[str] = []
        self._source_row_lookup: dict[str, int] = {}
        self._export_thread: QThread | None = None
        self._export_worker: _ExportWorker | None = None
        self._setup_ui()

    def _setup_ui(self) -> None:
        layout = QVBoxLayout(self)

        main_split = QSplitter(Qt.Orientation.Vertical)
        layout.addWidget(main_split, 1)

        # Haut: sources/détails + concat
        top_container = QWidget()
        top_row = QHBoxLayout(top_container)

        # Colonne gauche: sources + détails
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)

        source_group = QGroupBox("Colonnes source")
        source_layout = QVBoxLayout()
        source_header = QHBoxLayout()
        self._source_file_label = QLabel("Source: — (cliquez la liste pour charger)")
        self._source_sheet_combo = QComboBox()
        self._source_sheet_combo.currentTextChanged.connect(self._on_source_sheet_changed)
        self._source_header_spin = QSpinBox()
        self._source_header_spin.setRange(1, 1000000)
        self._source_header_spin.setValue(1)
        self._source_header_spin.valueChanged.connect(lambda _v: self._reload_source())
        source_header.addWidget(self._source_file_label, 1)
        source_header.addWidget(QLabel("Feuille:"))
        source_header.addWidget(self._source_sheet_combo)
        source_header.addWidget(QLabel("En-tête:"))
        source_header.addWidget(self._source_header_spin)
        source_layout.addLayout(source_header)

        source_tools = QHBoxLayout()
        source_tools.addWidget(QLabel("Apercu ligne:"))
        self._source_preview_row_spin = QSpinBox()
        self._source_preview_row_spin.setRange(0, 1000000)
        self._source_preview_row_spin.setValue(1)
        self._source_preview_row_spin.valueChanged.connect(self._on_source_preview_row_changed)
        source_tools.addWidget(self._source_preview_row_spin)
        source_tools.addSpacing(8)
        source_tools.addWidget(QLabel("Ordre:"))
        self._source_order_combo = QComboBox()
        self._source_order_combo.addItem("Origine", SOURCE_ORDER_ORIGIN)
        self._source_order_combo.addItem("A-Z", SOURCE_ORDER_AZ)
        self._source_order_combo.addItem("Z-A", SOURCE_ORDER_ZA)
        self._source_order_combo.addItem("Valeur", SOURCE_ORDER_VALUE)
        self._source_order_combo.addItem("Usage", SOURCE_ORDER_USAGE)
        self._source_order_combo.addItem("Manuel", SOURCE_ORDER_MANUAL)
        self._source_order_combo.currentIndexChanged.connect(self._on_source_order_changed)
        source_tools.addWidget(self._source_order_combo)
        source_tools.addStretch()
        source_layout.addLayout(source_tools)

        self._mapping_source_table = QTableWidget()
        self._mapping_source_table.setColumnCount(3)
        self._mapping_source_table.setHorizontalHeaderLabels(
            ["Colonne", "Apercu (ligne 1)", "Usage"]
        )
        self._mapping_source_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._mapping_source_table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self._mapping_source_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._mapping_source_table.verticalHeader().setVisible(False)
        src_header = self._mapping_source_table.horizontalHeader()
        src_header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        src_header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        src_header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self._mapping_source_table.setSortingEnabled(False)
        self._mapping_source_table.setDragDropOverwriteMode(False)
        self._mapping_source_table.setDragDropMode(QAbstractItemView.DragDropMode.NoDragDrop)
        self._mapping_source_table.setDragEnabled(False)
        self._mapping_source_table.setAcceptDrops(False)
        self._mapping_source_table.setDropIndicatorShown(False)
        self._mapping_source_table.viewport().installEventFilter(self)
        model = self._mapping_source_table.model()
        if model is not None:
            model.rowsMoved.connect(self._on_source_rows_moved)
        source_layout.addWidget(self._mapping_source_table, 1)
        source_group.setLayout(source_layout)
        left_layout.addWidget(source_group, 3)

        detail_group = QGroupBox("Détails")
        detail_layout = QVBoxLayout()
        header_row = QHBoxLayout()
        header_row.addWidget(QLabel("Cible:"))
        self._mapping_target_label = QLabel("—")
        self._mapping_target_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        header_row.addWidget(self._mapping_target_label, 1)
        self._mapping_mode_badge = QLabel("—")
        header_row.addWidget(self._mapping_mode_badge)
        detail_layout.addLayout(header_row)
        self._mapping_detail_hint = QLabel("Sélectionnez un champ cible dans la preview.")
        self._mapping_detail_hint.setWordWrap(True)
        detail_layout.addWidget(self._mapping_detail_hint)
        detail_group.setLayout(detail_layout)
        detail_group.setMinimumHeight(90)
        detail_group.setMaximumHeight(140)
        left_layout.addWidget(detail_group, 1)

        top_row.addWidget(left_panel, 3)

        # Colonne droite: concat
        concat_group = QGroupBox("Concaténation")
        concat_group_layout = QVBoxLayout()
        self._concat_empty_hint = QLabel("Sélectionnez un champ cible puis choisissez Concat.")
        self._concat_empty_hint.setWordWrap(True)
        concat_group_layout.addWidget(self._concat_empty_hint)

        self._mapping_concat_panel = QWidget()
        concat_layout = QVBoxLayout()
        concat_form = QFormLayout()
        self._concat_sep_edit = QLineEdit("; ")
        self._concat_sep_edit.setPlaceholderText("Ex: ; ou \\n")
        self._concat_sep_edit.editingFinished.connect(self._on_concat_changed)
        self._concat_skip_empty_cb = QCheckBox("Ignorer vides")
        self._concat_skip_empty_cb.setChecked(True)
        self._concat_skip_empty_cb.stateChanged.connect(lambda _v: self._on_concat_changed())
        self._concat_dedupe_cb = QCheckBox("Dédupliquer les valeurs")
        self._concat_dedupe_cb.stateChanged.connect(lambda _v: self._on_concat_changed())
        concat_form.addRow("Séparateur:", self._concat_sep_edit)
        concat_form.addRow("", self._concat_skip_empty_cb)
        concat_form.addRow("", self._concat_dedupe_cb)
        concat_layout.addLayout(concat_form)
        concat_layout.addWidget(QLabel("Colonnes concaténées (ordre)"))

        self._concat_sources_list = QListWidget()
        self._concat_sources_list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self._concat_sources_list.setMinimumHeight(200)
        self._concat_sources_list.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        concat_layout.addWidget(self._concat_sources_list, 1)

        concat_btns = QHBoxLayout()
        self._concat_add_btn = QPushButton("Ajouter colonne")
        self._concat_add_btn.clicked.connect(self._add_concat_source)
        self._concat_remove_btn = QPushButton("Supprimer")
        self._concat_remove_btn.clicked.connect(self._remove_concat_source)
        self._concat_up_btn = QPushButton("↑")
        self._concat_up_btn.clicked.connect(lambda: self._move_concat_source(-1))
        self._concat_down_btn = QPushButton("↓")
        self._concat_down_btn.clicked.connect(lambda: self._move_concat_source(1))
        concat_btns.addWidget(self._concat_add_btn)
        concat_btns.addWidget(self._concat_remove_btn)
        concat_btns.addStretch()
        concat_btns.addWidget(self._concat_up_btn)
        concat_btns.addWidget(self._concat_down_btn)
        concat_layout.addLayout(concat_btns)

        self._mapping_concat_panel.setLayout(concat_layout)
        concat_group_layout.addWidget(self._mapping_concat_panel, 1)
        concat_group.setLayout(concat_group_layout)
        concat_group.setMinimumHeight(260)

        top_row.addWidget(concat_group, 2)

        main_split.addWidget(top_container)

        # Bas: preview
        preview_group = QGroupBox("Prévisualisation")
        preview_layout = QVBoxLayout()

        template_controls = QHBoxLayout()
        self._template_file_label = QLabel("Template: — (cliquez la preview pour charger)")
        self._template_sheet_combo = QComboBox()
        self._template_sheet_combo.currentTextChanged.connect(self._on_template_sheet_changed)
        self._template_tech_row_spin = QSpinBox()
        self._template_tech_row_spin.setRange(1, 1000000)
        self._template_tech_row_spin.setValue(1)
        self._template_tech_row_spin.valueChanged.connect(lambda _v: self._refresh_mapping_preview())
        self._template_label_rows_edit = QLineEdit()
        self._template_label_rows_edit.setPlaceholderText("Lignes labels (ex: 1,2)")
        self._template_label_rows_edit.editingFinished.connect(self._refresh_mapping_preview)
        template_controls.addWidget(self._template_file_label, 1)
        template_controls.addWidget(QLabel("Feuille:"))
        template_controls.addWidget(self._template_sheet_combo)
        template_controls.addWidget(QLabel("Champs cible ligne:"))
        template_controls.addWidget(self._template_tech_row_spin)
        template_controls.addWidget(QLabel("Labels:"))
        template_controls.addWidget(self._template_label_rows_edit)
        preview_layout.addLayout(template_controls)

        preview_controls = QHBoxLayout()
        preview_controls.addWidget(QLabel("Lignes de données:"))
        self._mapping_preview_rows_spin = QSpinBox()
        self._mapping_preview_rows_spin.setRange(1, 1000000)
        self._mapping_preview_rows_spin.setValue(5)
        self._mapping_preview_rows_spin.valueChanged.connect(lambda _v: self._refresh_mapping_preview())
        preview_controls.addWidget(self._mapping_preview_rows_spin)
        self._mapping_preview_refresh_btn = QPushButton("Rafraîchir")
        self._mapping_preview_refresh_btn.clicked.connect(self._refresh_mapping_preview)
        preview_controls.addWidget(self._mapping_preview_refresh_btn)
        preview_controls.addStretch()
        preview_layout.addLayout(preview_controls)

        self._mapping_preview_label = QLabel("")
        preview_layout.addWidget(self._mapping_preview_label)

        self._mapping_preview_table = QTableWidget()
        self._mapping_preview_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self._mapping_preview_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self._mapping_preview_table.verticalHeader().setVisible(False)
        header = self._mapping_preview_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(False)
        self._mapping_preview_table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self._mapping_preview_table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self._mapping_preview_table.cellClicked.connect(self._on_preview_cell_clicked)
        self._mapping_preview_table.viewport().installEventFilter(self)
        preview_layout.addWidget(self._mapping_preview_table, 1)

        preview_group.setLayout(preview_layout)
        main_split.addWidget(preview_group)

        main_split.setStretchFactor(0, 3)
        main_split.setStretchFactor(1, 2)
        main_split.setChildrenCollapsible(False)
        main_split.setSizes([600, 400])

        # Footer
        footer = QHBoxLayout()
        footer.addStretch()
        self._export_btn = QPushButton("Exporter")
        self._export_btn.clicked.connect(self._export)
        self._export_menu = QMenu(self._export_btn)
        export_all = self._export_menu.addAction("Exporter (xlsx/ods/csv)")
        export_all.triggered.connect(lambda _checked=False: self._export())
        export_csv_full = self._export_menu.addAction("Exporter CSV (complet)")
        export_csv_full.triggered.connect(lambda _checked=False: self._export_csv(False))
        export_csv_trim = self._export_menu.addAction("Exporter CSV (sans colonnes vides)")
        export_csv_trim.triggered.connect(lambda _checked=False: self._export_csv(True))
        self._export_btn.setMenu(self._export_menu)
        footer.addWidget(self._export_btn)
        layout.addLayout(footer)

        self._mapping_concat_panel.setVisible(False)

    def eventFilter(self, obj: QObject, event: QEvent) -> bool:  # type: ignore[name-defined]
        if event.type() == QEvent.Type.MouseButtonPress:
            if obj == self._mapping_preview_table.viewport() and not self._template_path:
                self._browse_template()
                return True
            if obj == self._mapping_source_table.viewport() and not self._source_path:
                self._browse_source()
                return True
        return super().eventFilter(obj, event)

    # --- Loading template/source ---
    def _browse_template(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Choisir un template", "", SUPPORTED_INPUT_FILTER)
        if not path:
            return
        self._template_path = path
        self._template_file_label.setText(f"Template: {Path(path).name}")
        self._load_template_sheets()

    def _browse_source(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Choisir une source", "", SUPPORTED_INPUT_FILTER)
        if not path:
            return
        self._source_path = path
        self._source_file_label.setText(f"Source: {Path(path).name}")
        self._load_source_sheets()

    def _load_template_sheets(self) -> None:
        try:
            sheets = list_sheets(self._template_path)
        except ExcelFileError as e:
            QMessageBox.critical(self, "Erreur", str(e))
            return
        self._template_sheet_combo.blockSignals(True)
        self._template_sheet_combo.clear()
        self._template_sheet_combo.addItems(sheets)
        self._template_sheet_combo.blockSignals(False)
        if sheets:
            self._template_sheet_combo.setCurrentIndex(0)
            self._reload_template()

    def _load_source_sheets(self) -> None:
        try:
            sheets = list_sheets(self._source_path)
        except ExcelFileError as e:
            QMessageBox.critical(self, "Erreur", str(e))
            return
        self._source_sheet_combo.blockSignals(True)
        self._source_sheet_combo.clear()
        self._source_sheet_combo.addItems(sheets)
        self._source_sheet_combo.blockSignals(False)
        if sheets:
            self._source_sheet_combo.setCurrentIndex(0)
            self._reload_source()

    def _on_template_sheet_changed(self, _sheet: str) -> None:
        if self._template_path:
            self._reload_template()

    def _on_source_sheet_changed(self, _sheet: str) -> None:
        if self._source_path:
            self._reload_source()

    def _reload_template(self) -> None:
        if not self._template_path:
            return
        sheet = self._template_sheet_combo.currentText() or None
        try:
            self._template_df_raw = load_sheet_raw(self._template_path, sheet)
        except ExcelFileError as e:
            QMessageBox.critical(self, "Erreur", str(e))
            return
        self._update_zone_bounds()
        self._refresh_mapping_preview()

    def _reload_source(self) -> None:
        if not self._source_path:
            return
        sheet = self._source_sheet_combo.currentText() or None
        try:
            self._source_df = load_sheet(self._source_path, sheet, header_row=self._source_header_spin.value())
        except ExcelFileError as e:
            QMessageBox.critical(self, "Erreur", str(e))
            return
        self._refresh_mapping_sources()
        self._refresh_mapping_preview()

    # --- Zone/mapping helpers ---
    def _update_zone_bounds(self) -> None:
        if self._template_df_raw is None:
            return
        self._zone["row_start"] = 1
        self._zone["col_start"] = 1
        self._zone["row_end"] = len(self._template_df_raw)
        self._zone["col_end"] = self._template_df_raw.shape[1]

    def _update_zone_header(self) -> None:
        header = {
            "tech_row": self._template_tech_row_spin.value(),
            "label_rows": _parse_int_list(self._template_label_rows_edit.text()),
        }
        self._zone["header"] = header

    def _get_source_cols(self) -> list[str]:
        return list(self._source_df.columns) if self._source_df is not None else []

    def _refresh_mapping_sources(self) -> None:
        source_cols = self._get_source_cols()
        self._source_default_order = list(source_cols)
        self._source_manual_order = list(source_cols)
        self._source_row_lookup = {}
        self._source_order_mode = SOURCE_ORDER_ORIGIN
        self._source_order_combo.blockSignals(True)
        idx = self._source_order_combo.findData(SOURCE_ORDER_ORIGIN)
        if idx >= 0:
            self._source_order_combo.setCurrentIndex(idx)
        self._source_order_combo.blockSignals(False)
        self._apply_source_order_mode()
        self._refresh_source_table(keep_selection=False)
        self._refresh_concat_source_widgets(source_cols)

    def _refresh_concat_source_widgets(self, source_cols: list[str]) -> None:
        for i in range(self._concat_sources_list.count()):
            item = self._concat_sources_list.item(i)
            widget = self._concat_sources_list.itemWidget(item)
            if isinstance(widget, _ConcatSourceWidget):
                widget.refresh_source_cols(source_cols)

    def _on_source_preview_row_changed(self) -> None:
        self._refresh_source_table(keep_selection=True)

    def _on_source_order_changed(self) -> None:
        mode = self._source_order_combo.currentData()
        if mode not in (
            SOURCE_ORDER_ORIGIN,
            SOURCE_ORDER_AZ,
            SOURCE_ORDER_ZA,
            SOURCE_ORDER_VALUE,
            SOURCE_ORDER_USAGE,
            SOURCE_ORDER_MANUAL,
        ):
            return
        if mode == SOURCE_ORDER_MANUAL and not self._source_manual_order:
            self._source_manual_order = self._get_source_table_order() or list(self._source_default_order)
        self._source_order_mode = mode
        self._apply_source_order_mode()
        self._refresh_source_table(keep_selection=True)

    def _apply_source_order_mode(self) -> None:
        if self._source_order_mode == SOURCE_ORDER_MANUAL:
            self._mapping_source_table.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
            self._mapping_source_table.setDragEnabled(True)
            self._mapping_source_table.setAcceptDrops(True)
            self._mapping_source_table.setDropIndicatorShown(True)
            self._mapping_source_table.setDefaultDropAction(Qt.DropAction.MoveAction)
        else:
            self._mapping_source_table.setDragDropMode(QAbstractItemView.DragDropMode.NoDragDrop)
            self._mapping_source_table.setDragEnabled(False)
            self._mapping_source_table.setAcceptDrops(False)
            self._mapping_source_table.setDropIndicatorShown(False)

    def _on_source_rows_moved(self, *_args: object) -> None:
        if self._source_order_mode != SOURCE_ORDER_MANUAL:
            return
        order = self._get_source_table_order()
        self._source_manual_order = order
        self._source_row_lookup = {name: idx for idx, name in enumerate(order)}

    def _get_source_table_selected_names(self) -> list[str]:
        model = self._mapping_source_table.selectionModel()
        if model is None:
            return []
        names: list[str] = []
        for index in model.selectedRows(0):
            item = self._mapping_source_table.item(index.row(), 0)
            if item is None:
                continue
            names.append(item.data(Qt.ItemDataRole.UserRole) or item.text())
        return names

    def _set_source_table_selection(self, names: list[str]) -> None:
        model = self._mapping_source_table.selectionModel()
        if model is None:
            return
        self._mapping_source_table.clearSelection()
        if not names:
            return
        for name in names:
            row = self._source_row_lookup.get(name)
            if row is None:
                continue
            index = self._mapping_source_table.model().index(row, 0)
            model.select(index, QItemSelectionModel.SelectionFlag.Select | QItemSelectionModel.SelectionFlag.Rows)
        first_row = self._source_row_lookup.get(names[0])
        if first_row is not None:
            self._mapping_source_table.setCurrentCell(first_row, 0)
            item = self._mapping_source_table.item(first_row, 0)
            if item is not None:
                self._mapping_source_table.scrollToItem(
                    item, QAbstractItemView.ScrollHint.PositionAtCenter
                )

    def _get_source_table_order(self) -> list[str]:
        order: list[str] = []
        for row in range(self._mapping_source_table.rowCount()):
            item = self._mapping_source_table.item(row, 0)
            if item is None:
                continue
            order.append(item.data(Qt.ItemDataRole.UserRole) or item.text())
        return order

    def _compute_source_usage_counts(self) -> dict[str, int]:
        counts: dict[str, int] = {}
        for mapping in self._zone.get("field_mappings", []):
            if mapping.get("mode") == "simple":
                src = mapping.get("source_col") or ""
                if src:
                    counts[src] = counts.get(src, 0) + 1
            elif mapping.get("mode") == "concat":
                concat = mapping.get("concat") or {}
                for src in concat.get("sources", []):
                    col = src.get("col") or ""
                    if col:
                        counts[col] = counts.get(col, 0) + 1
        return counts

    def _compute_source_preview_values(self, source_cols: list[str]) -> tuple[dict[str, str], bool]:
        values: dict[str, str] = {}
        row_value = self._source_preview_row_spin.value()
        if row_value == 0:
            for col in source_cols:
                values[col] = col
            return values, False
        if self._source_df is None:
            return values, False
        row_index = row_value - 1
        if row_index < 0 or row_index >= len(self._source_df):
            return values, True
        for idx, col in enumerate(source_cols):
            if idx >= self._source_df.shape[1]:
                continue
            val = self._source_df.iat[row_index, idx]
            values[col] = _format_value(val)
        return values, False

    def _apply_source_order(
        self,
        source_cols: list[str],
        preview_values: dict[str, str],
        usage_counts: dict[str, int],
    ) -> list[str]:
        mode = self._source_order_mode
        if mode == SOURCE_ORDER_AZ:
            return sorted(source_cols, key=lambda x: str(x).lower())
        if mode == SOURCE_ORDER_ZA:
            return sorted(source_cols, key=lambda x: str(x).lower(), reverse=True)
        if mode == SOURCE_ORDER_VALUE:
            return sorted(
                source_cols,
                key=lambda x: (
                    preview_values.get(x, "") == "",
                    preview_values.get(x, "").lower(),
                ),
            )
        if mode == SOURCE_ORDER_USAGE:
            return sorted(
                source_cols,
                key=lambda x: -usage_counts.get(x, 0),
            )
        if mode == SOURCE_ORDER_MANUAL:
            ordered = [col for col in self._source_manual_order if col in source_cols]
            for col in source_cols:
                if col not in ordered:
                    ordered.append(col)
            return ordered
        return list(source_cols)

    def _refresh_source_table(self, *, keep_selection: bool) -> None:
        source_cols = self._get_source_cols()
        if not source_cols:
            self._mapping_source_table.clearContents()
            self._mapping_source_table.setRowCount(0)
            self._source_row_lookup = {}
            return
        selected = self._get_source_table_selected_names() if keep_selection else []
        usage_counts = self._compute_source_usage_counts()
        preview_values, out_of_range = self._compute_source_preview_values(source_cols)
        ordered_cols = self._apply_source_order(source_cols, preview_values, usage_counts)

        self._mapping_source_table.setRowCount(len(ordered_cols))
        self._source_row_lookup = {}
        for row_idx, col in enumerate(ordered_cols):
            self._source_row_lookup[col] = row_idx

            item_col = QTableWidgetItem(col)
            item_col.setData(Qt.ItemDataRole.UserRole, col)
            item_col.setFlags(item_col.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self._mapping_source_table.setItem(row_idx, 0, item_col)

            raw_value = preview_values.get(col, "")
            if out_of_range:
                display_value = "(hors limites)"
                placeholder = True
            elif raw_value == "":
                display_value = "(vide)"
                placeholder = True
            else:
                display_value = raw_value
                placeholder = False
            item_val = QTableWidgetItem(display_value)
            item_val.setData(Qt.ItemDataRole.UserRole, raw_value)
            item_val.setToolTip(raw_value)
            item_val.setFlags(item_val.flags() & ~Qt.ItemFlag.ItemIsEditable)
            if placeholder:
                item_val.setForeground(Qt.GlobalColor.gray)
            self._mapping_source_table.setItem(row_idx, 1, item_val)

            count = usage_counts.get(col, 0)
            item_usage = QTableWidgetItem("" if count == 0 else str(count))
            item_usage.setData(Qt.ItemDataRole.UserRole, count)
            item_usage.setTextAlignment(
                Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter
            )
            item_usage.setFlags(item_usage.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self._mapping_source_table.setItem(row_idx, 2, item_usage)

        header_item = self._mapping_source_table.horizontalHeaderItem(1)
        if header_item is not None:
            header_item.setText(f"Apercu (ligne {self._source_preview_row_spin.value()})")

        if keep_selection:
            self._set_source_table_selection(selected)
        else:
            self._mapping_source_table.clearSelection()

    def _get_zone_target_columns(self) -> list[dict[str, Any]]:
        if self._template_df_raw is None:
            return []
        df = self._template_df_raw
        row_start = int(self._zone.get("row_start", 1)) - 1
        row_end = int(self._zone.get("row_end", len(df))) - 1
        col_start = int(self._zone.get("col_start", 1)) - 1
        col_end = int(self._zone.get("col_end", df.shape[1])) - 1
        header = self._zone.get("header", {})
        tech_row = int(header.get("tech_row", row_start + 1)) - 1
        tech_row = max(0, min(tech_row, len(df) - 1))
        labels = df.iloc[tech_row, col_start : col_end + 1].tolist()
        targets = []
        for idx, label in enumerate(labels):
            text = "" if label is None else str(label)
            targets.append({"col_index": idx, "label": text})
        return targets

    def _current_mapping_target(self) -> tuple[dict[str, Any], int, str, int] | None:
        if not self._current_mapping_cols:
            return None
        col_idx = self._current_preview_col if self._current_preview_col is not None else 0
        if col_idx < 0 or col_idx >= len(self._current_mapping_cols):
            return None
        label = self._current_mapping_labels[col_idx]
        col_index = self._current_mapping_cols[col_idx]
        return self._zone, col_idx, label, col_index

    def _set_mapping(self, zone: dict[str, Any], target: str, col_index: int, data: dict[str, Any]) -> None:
        mappings = list(zone.get("field_mappings", []))
        for i, m in enumerate(mappings):
            if target and m.get("target") == target:
                mappings[i] = data
                zone["field_mappings"] = mappings
                return
        for i, m in enumerate(mappings):
            if m.get("col_index") == col_index:
                mappings[i] = data
                zone["field_mappings"] = mappings
                return
        mappings.append(data)
        zone["field_mappings"] = mappings

    def _remove_mapping(self, zone: dict[str, Any], target: str, col_index: int) -> None:
        mappings = []
        for m in zone.get("field_mappings", []):
            if target and m.get("target") == target:
                continue
            if m.get("col_index") == col_index:
                continue
            mappings.append(m)
        zone["field_mappings"] = mappings

    def _set_mapping_mode_badge(self, mode: str) -> None:
        labels = {"ignore": "Ignorer", "simple": "Simple", "concat": "Concat"}
        self._mapping_mode_badge.setText(labels.get(mode, "—"))

    def _select_target_col(self, col_idx: int) -> None:
        self._current_preview_col = col_idx
        try:
            self._mapping_preview_table.setCurrentCell(0, col_idx)
        except Exception:
            pass
        info = self._current_mapping_target()
        if info is None:
            self._clear_mapping_detail()
            return
        _, _, label, _ = info
        mapping = self._get_mapping(label)
        mode = mapping.get("mode") if mapping else "ignore"
        self._mapping_target_label.setText(label or "(sans label)")
        self._set_mapping_mode_badge(mode)
        if mode == "simple":
            self._mapping_detail_hint.setText("Mapping simple défini via la preview.")
            self._mapping_detail_hint.setVisible(True)
            self._mapping_concat_panel.setVisible(False)
            self._concat_empty_hint.setVisible(True)
        elif mode == "concat":
            self._mapping_detail_hint.setVisible(False)
            self._mapping_concat_panel.setVisible(True)
            self._concat_empty_hint.setVisible(False)
            self._load_concat_editor(mapping.get("concat") if mapping else None)
        else:
            self._mapping_detail_hint.setText("Aucune association (ignore).")
            self._mapping_detail_hint.setVisible(True)
            self._mapping_concat_panel.setVisible(False)
            self._concat_empty_hint.setVisible(True)
        self._sync_source_selection(mapping)

    def _clear_mapping_detail(self) -> None:
        self._mapping_target_label.setText("—")
        self._mapping_mode_badge.setText("—")
        self._mapping_detail_hint.setText("Sélectionnez un champ cible dans la preview.")
        self._mapping_detail_hint.setVisible(True)
        self._mapping_concat_panel.setVisible(False)
        self._concat_empty_hint.setVisible(True)

    def _sync_source_selection(self, mapping: dict[str, Any] | None) -> None:
        if not mapping:
            return
        mode = mapping.get("mode")
        cols: list[str] = []
        if mode == "simple":
            src = mapping.get("source_col") or ""
            if src:
                cols = [src]
        elif mode == "concat":
            concat = mapping.get("concat") or {}
            for src in concat.get("sources", []):
                col = src.get("col") or ""
                if col:
                    cols.append(col)
        if cols:
            self._set_source_table_selection(cols)

    def _get_mapping(self, target: str) -> dict[str, Any] | None:
        for m in self._zone.get("field_mappings", []):
            if target and m.get("target") == target:
                return m
        return None

    def _get_primary_source_column(self) -> str:
        row = self._mapping_source_table.currentRow()
        if row >= 0:
            item = self._mapping_source_table.item(row, 0)
            if item is not None:
                return item.data(Qt.ItemDataRole.UserRole) or item.text()
        selected = self._get_source_table_selected_names()
        return selected[0] if selected else ""

    def _selected_source_columns(self) -> list[str]:
        return self._get_source_table_selected_names()

    # --- Preview ---
    def _calc_header_rows(self, zone: dict[str, Any]) -> int:
        row_start = int(zone.get("row_start", 1))
        header = zone.get("header", {})
        rows = []
        rows.extend([int(x) for x in header.get("label_rows", [])])
        tech_row = header.get("tech_row")
        if tech_row:
            rows.append(int(tech_row))
        header_end = max(rows) if rows else row_start
        return max(0, header_end - row_start + 1)

    def _refresh_mapping_preview(self) -> None:
        if self._template_df_raw is None or self._source_df is None:
            self._mapping_preview_table.clear()
            self._mapping_preview_label.setText("Chargez un template et une source.")
            return
        self._update_zone_header()
        max_rows = self._mapping_preview_rows_spin.value()
        zone_spec = ZoneSpec.from_dict(self._zone)
        try:
            df = _build_zone_output(zone_spec, self._template_df_raw, self._source_df.head(max_rows))
        except Exception as e:
            self._mapping_preview_table.clear()
            self._mapping_preview_label.setText(f"Erreur preview: {e}")
            return

        targets = self._get_zone_target_columns()
        self._current_mapping_cols = [t["col_index"] for t in targets]
        self._current_mapping_labels = [t["label"] for t in targets]
        col_count = len(targets)

        header_rows = self._calc_header_rows(self._zone)
        data_df = df
        if header_rows > 0:
            data_df = df.iloc[header_rows:, :].copy()
        if len(data_df) > max_rows:
            data_df = data_df.iloc[:max_rows, :].copy()

        row_count = 1 + len(data_df)
        self._mapping_preview_table.clear()
        self._mapping_preview_table.setRowCount(row_count)
        self._mapping_preview_table.setColumnCount(col_count)
        self._mapping_preview_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)

        self._preview_combo_by_col = {}
        self._preview_concat_btn_by_col = {}
        source_cols = self._get_source_cols()

        for col_idx, target in enumerate(targets):
            label = target["label"]
            mapping = self._get_mapping(label)
            cell = self._build_mapping_cell_widget(col_idx, label, mapping, source_cols)
            self._mapping_preview_table.setCellWidget(0, col_idx, cell)

        for r in range(len(data_df)):
            for c in range(col_count):
                value = ""
                if c < data_df.shape[1]:
                    val = data_df.iat[r, c]
                    value = "" if pd.isna(val) else str(val)
                item = QTableWidgetItem(value)
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self._mapping_preview_table.setItem(r + 1, c, item)

        self._mapping_preview_table.setRowHeight(0, 64)
        self._adjust_preview_column_widths()

        if self._current_preview_col is None and col_count > 0:
            self._current_preview_col = 0
        if self._current_preview_col is not None and self._current_preview_col < col_count:
            self._select_target_col(self._current_preview_col)
        self._refresh_source_usage()

        self._mapping_preview_label.setText(f"{len(data_df)} lignes × {col_count} colonnes")

    def _build_mapping_cell_widget(
        self,
        col_idx: int,
        label: str,
        mapping: dict[str, Any] | None,
        source_cols: list[str],
    ) -> QWidget:
        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(4, 2, 4, 2)
        layout.setSpacing(4)

        title = QLabel(label or "(sans label)")
        title.setWordWrap(False)
        layout.addWidget(title)

        row = QHBoxLayout()
        combo = QComboBox()
        combo.addItem("—", "")
        combo.addItem("Concat…", CONCAT_MENU_VALUE)
        for col in source_cols:
            combo.addItem(col, col)
        combo.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToContents)

        current_value = ""
        if mapping:
            mode = mapping.get("mode")
            if mode == "concat":
                current_value = CONCAT_MENU_VALUE
            elif mode == "simple":
                current_value = mapping.get("source_col") or ""
        if current_value:
            idx = combo.findData(current_value)
            if idx >= 0:
                combo.blockSignals(True)
                combo.setCurrentIndex(idx)
                combo.blockSignals(False)

        combo.currentTextChanged.connect(lambda _text, c=col_idx: self._on_preview_combo_changed(c))
        row.addWidget(combo, 1)

        concat_btn = QToolButton()
        concat_btn.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView))
        concat_btn.setToolTip("Concat")
        concat_btn.setFixedWidth(26)
        concat_btn.clicked.connect(lambda _=None, c=col_idx: self._on_preview_concat_clicked(c))
        row.addWidget(concat_btn)
        layout.addLayout(row)

        self._preview_combo_by_col[col_idx] = combo
        self._preview_concat_btn_by_col[col_idx] = concat_btn
        return container

    def _adjust_preview_column_widths(self) -> None:
        table = self._mapping_preview_table
        min_width = 160
        for col in range(table.columnCount()):
            width = table.sizeHintForColumn(col)
            widget = table.cellWidget(0, col)
            if widget is not None:
                width = max(width, widget.sizeHint().width())
            table.setColumnWidth(col, max(min_width, width + 12))

    def _on_preview_cell_clicked(self, row: int, col: int) -> None:
        if row < 0 or col < 0:
            return
        self._select_target_col(col)

    def _on_preview_combo_changed(self, col_idx: int) -> None:
        combo = self._preview_combo_by_col.get(col_idx)
        if combo is None:
            return
        value = combo.currentData()
        self._select_target_col(col_idx)
        if value == CONCAT_MENU_VALUE:
            self._map_concat_current()
            return
        if value in ("", None):
            self._remove_mapping_current()
            return
        info = self._current_mapping_target()
        if info is None:
            return
        zone, _, label, col_index = info
        data = {
            "col_index": col_index,
            "target": label,
            "mode": "simple",
            "source_col": str(value),
        }
        self._set_mapping(zone, label, col_index, data)
        self._set_mapping_mode_badge("simple")
        self._refresh_mapping_preview()
        self._refresh_source_usage()

    def _on_preview_concat_clicked(self, col_idx: int) -> None:
        self._select_target_col(col_idx)
        self._map_concat_current()

    # --- Mapping actions ---
    def _map_concat_current(self) -> None:
        info = self._current_mapping_target()
        if info is None:
            QMessageBox.warning(self, "Attention", "Sélectionnez un champ cible.")
            return
        zone, _, label, col_index = info
        mapping = self._get_mapping(label)
        concat = mapping.get("concat") if mapping and mapping.get("mode") == "concat" else None
        if concat is None:
            concat = {"separator": "; ", "skip_empty": True, "deduplicate": False, "sources": []}
        selected_sources = self._selected_source_columns()
        if selected_sources:
            existing = [src.get("col") for src in concat.get("sources", [])]
            for src in selected_sources:
                if src not in existing:
                    concat.setdefault("sources", []).append({"col": src, "prefix": ""})
        data = {
            "col_index": col_index,
            "target": label,
            "mode": "concat",
            "concat": concat,
        }
        self._set_mapping(zone, label, col_index, data)
        self._set_mapping_mode_badge("concat")
        self._refresh_mapping_preview()
        self._refresh_source_usage()

    def _remove_mapping_current(self) -> None:
        info = self._current_mapping_target()
        if info is None:
            return
        zone, _, label, col_index = info
        self._remove_mapping(zone, label, col_index)
        self._clear_mapping_detail()
        self._refresh_mapping_preview()
        self._refresh_source_usage()

    # --- Concat editor ---
    def _add_concat_source(self, col: str | None = None, prefix: str = "") -> None:
        source_cols = self._get_source_cols()
        default_col = col if col is not None else self._get_primary_source_column()
        widget = _ConcatSourceWidget(source_cols, col=default_col, prefix=prefix)
        widget.changed.connect(self._on_concat_changed)
        item = QListWidgetItem()
        item.setSizeHint(widget.sizeHint())
        self._concat_sources_list.addItem(item)
        self._concat_sources_list.setItemWidget(item, widget)
        self._concat_sources_list.setCurrentItem(item)
        self._on_concat_changed()

    def _remove_concat_source(self) -> None:
        row = self._concat_sources_list.currentRow()
        if row < 0:
            return
        item = self._concat_sources_list.item(row)
        widget = self._concat_sources_list.itemWidget(item) if item else None
        self._concat_sources_list.takeItem(row)
        if widget:
            widget.deleteLater()
        self._on_concat_changed()

    def _move_concat_source(self, delta: int) -> None:
        row = self._concat_sources_list.currentRow()
        if row < 0:
            return
        new_row = row + delta
        if new_row < 0 or new_row >= self._concat_sources_list.count():
            return
        item = self._concat_sources_list.item(row)
        widget = self._concat_sources_list.itemWidget(item) if item else None
        self._concat_sources_list.takeItem(row)
        if widget:
            widget.setParent(None)
        self._concat_sources_list.insertItem(new_row, item)
        if widget:
            self._concat_sources_list.setItemWidget(item, widget)
        self._concat_sources_list.setCurrentRow(new_row)
        self._on_concat_changed()

    def _load_concat_editor(self, concat: dict[str, Any] | None) -> None:
        self._mapping_concat_loading = True
        data = concat or {}
        self._concat_sep_edit.setText(str(data.get("separator", "; ")))
        self._concat_skip_empty_cb.setChecked(bool(data.get("skip_empty", True)))
        self._concat_dedupe_cb.setChecked(bool(data.get("deduplicate", False)))
        self._concat_sources_list.clear()
        for src in data.get("sources", []):
            self._add_concat_source(col=src.get("col", ""), prefix=src.get("prefix", ""))
        self._mapping_concat_loading = False

    def _collect_concat_editor_data(self) -> dict[str, Any]:
        sources: list[dict[str, str]] = []
        for i in range(self._concat_sources_list.count()):
            item = self._concat_sources_list.item(i)
            widget = self._concat_sources_list.itemWidget(item)
            if isinstance(widget, _ConcatSourceWidget):
                col, prefix = widget.get_data()
                if col:
                    sources.append({"col": col, "prefix": prefix})
        return {
            "separator": self._concat_sep_edit.text(),
            "skip_empty": self._concat_skip_empty_cb.isChecked(),
            "deduplicate": self._concat_dedupe_cb.isChecked(),
            "sources": sources,
        }

    def _on_concat_changed(self) -> None:
        if self._mapping_concat_loading:
            return
        info = self._current_mapping_target()
        if info is None:
            return
        zone, _, label, col_index = info
        mapping = self._get_mapping(label)
        if not mapping or mapping.get("mode") != "concat":
            return
        concat = self._collect_concat_editor_data()
        data = {
            "col_index": col_index,
            "target": label,
            "mode": "concat",
            "concat": concat,
        }
        self._set_mapping(zone, label, col_index, data)
        self._refresh_mapping_preview()
        self._refresh_source_usage()

    # --- Source usage ---
    def _refresh_source_usage(self) -> None:
        self._refresh_source_table(keep_selection=True)

    # --- Export ---
    def _build_export_config(self) -> TemplateBuilderConfig:
        self._update_zone_header()
        zone_spec = ZoneSpec.from_dict(self._zone)
        return TemplateBuilderConfig(
            template_file=self._template_path,
            template_sheet=self._template_sheet_combo.currentText() or None,
            source_file=self._source_path,
            source_sheet=self._source_sheet_combo.currentText() or None,
            source_header_row=self._source_header_spin.value(),
            zones=[zone_spec],
            output_sheet_name="Output",
        )

    def _export(self) -> None:
        if not self._template_path or not self._source_path:
            QMessageBox.warning(self, "Attention", "Chargez un template et une source.")
            return
        if self._export_thread is not None and self._export_thread.isRunning():
            QMessageBox.information(self, "Export", "Un export est deja en cours.")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Exporter", "", SUPPORTED_OUTPUT_FILTER)
        if not path:
            return
        try:
            config = self._build_export_config()
            self._start_export(config, path, csv_separator=";", drop_empty_columns=False)
        except Exception as e:
            QMessageBox.critical(self, "Erreur", str(e))

    def _export_csv(self, drop_empty_columns: bool) -> None:
        if not self._template_path or not self._source_path:
            QMessageBox.warning(self, "Attention", "Chargez un template et une source.")
            return
        if self._export_thread is not None and self._export_thread.isRunning():
            QMessageBox.information(self, "Export", "Un export est deja en cours.")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Exporter CSV", "", "CSV (*.csv)")
        if not path:
            return
        if Path(path).suffix.lower() != ".csv":
            path = f"{path}.csv"
        try:
            config = self._build_export_config()
            self._start_export(config, path, csv_separator=";", drop_empty_columns=drop_empty_columns)
        except Exception as e:
            QMessageBox.critical(self, "Erreur", str(e))

    def _start_export(
        self,
        config: TemplateBuilderConfig,
        path: str,
        *,
        csv_separator: str = ";",
        drop_empty_columns: bool = False,
    ) -> None:
        self._export_btn.setEnabled(False)
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        self._export_thread = QThread(self)
        self._export_worker = _ExportWorker(
            config,
            path,
            csv_separator=csv_separator,
            drop_empty_columns=drop_empty_columns,
        )
        self._export_worker.moveToThread(self._export_thread)
        self._export_thread.started.connect(self._export_worker.run)
        self._export_worker.finished.connect(self._on_export_finished)
        self._export_worker.finished.connect(self._export_thread.quit)
        self._export_worker.finished.connect(self._export_worker.deleteLater)
        self._export_thread.finished.connect(self._export_thread.deleteLater)
        self._export_thread.start()

    def _on_export_finished(self, ok: bool, message: str) -> None:
        QApplication.restoreOverrideCursor()
        self._export_btn.setEnabled(True)
        self._export_worker = None
        self._export_thread = None
        if ok:
            QMessageBox.information(self, "Export", "Export termine.")
        else:
            QMessageBox.critical(self, "Erreur", message)

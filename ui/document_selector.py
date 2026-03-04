# =============================================================================
#  ui/document_selector.py  –  Dialogo selezione documento (per BOM ecc.)
# =============================================================================
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
    QComboBox
)
from PyQt6.QtCore import Qt

from ui.session import session
from ui.styles import TYPE_ICON


class DocumentSelectorDialog(QDialog):
    def __init__(self, parent=None, only_type: str = ""):
        super().__init__(parent)
        self.setWindowTitle("Seleziona Documento")
        self.setMinimumSize(640, 400)
        self.selected_id   = None
        self.selected_code = None
        self._only_type    = only_type
        self._build_ui()
        self._search()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(8)
        layout.setContentsMargins(12, 12, 12, 12)

        # Ricerca
        row = QHBoxLayout()
        self.txt_search = QLineEdit()
        self.txt_search.setPlaceholderText("Cerca per codice o titolo…")
        self.txt_search.returnPressed.connect(self._search)
        self.cmb_type   = QComboBox()
        self.cmb_type.addItems(["Tutti", "Parte", "Assieme", "Disegno"])
        if self._only_type:
            idx = self.cmb_type.findText(self._only_type)
            if idx >= 0:
                self.cmb_type.setCurrentIndex(idx)
            self.cmb_type.setEnabled(False)
        btn_search = QPushButton("Cerca")
        btn_search.clicked.connect(self._search)
        row.addWidget(self.txt_search)
        row.addWidget(self.cmb_type)
        row.addWidget(btn_search)
        layout.addLayout(row)

        # Tabella
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["Codice", "Rev.", "Tipo", "Titolo"])
        self.table.horizontalHeader().setSectionResizeMode(
            3, QHeaderView.ResizeMode.Stretch
        )
        self.table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows
        )
        self.table.setEditTriggers(
            QTableWidget.EditTrigger.NoEditTriggers
        )
        self.table.doubleClicked.connect(self._select)
        layout.addWidget(self.table)

        # Bottoni
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_ok = QPushButton("Seleziona")
        btn_ok.setObjectName("btn_primary")
        btn_ok.clicked.connect(self._select)
        btn_cancel = QPushButton("Annulla")
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_ok)
        layout.addLayout(btn_row)

    def _search(self):
        text     = self.txt_search.text().strip()
        doc_type = self.cmb_type.currentText()
        if doc_type == "Tutti":
            doc_type = ""
        docs = session.files.search_documents(text=text, doc_type=doc_type)
        self.table.setRowCount(0)
        for doc in docs:
            r = self.table.rowCount()
            self.table.insertRow(r)
            self.table.setItem(r, 0, QTableWidgetItem(doc["code"]))
            self.table.setItem(r, 1, QTableWidgetItem(doc["revision"]))
            self.table.setItem(
                r, 2, QTableWidgetItem(
                    TYPE_ICON.get(doc["doc_type"], "") + " " + doc["doc_type"]
                )
            )
            self.table.setItem(r, 3, QTableWidgetItem(doc["title"]))
            self.table.item(r, 0).setData(Qt.ItemDataRole.UserRole, doc["id"])

    def _select(self):
        row = self.table.currentRow()
        if row < 0:
            return
        item = self.table.item(row, 0)
        self.selected_id   = item.data(Qt.ItemDataRole.UserRole)
        self.selected_code = item.text()
        self.accept()

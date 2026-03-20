# =============================================================================
#  ui/commercial_item_selector.py  –  Selettore articolo commerciale per BOM
# =============================================================================
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
    QComboBox, QDialogButtonBox,
)
from PyQt6.QtCore import Qt

from ui.session import session


class CommercialItemSelectorDialog(QDialog):
    """Dialog per selezionare un articolo commerciale da inserire in una BOM CAD."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Seleziona articolo commerciale")
        self.setMinimumSize(700, 450)
        self.selected_id: int | None = None
        self._build_ui()
        self._refresh()

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setSpacing(8)
        lay.setContentsMargins(12, 12, 12, 12)

        # Filtri
        filter_row = QHBoxLayout()
        filter_row.addWidget(QLabel("Tipo:"))
        self.cmb_type = QComboBox()
        self.cmb_type.addItem("Tutti", "")
        self.cmb_type.addItem("Commerciali (5)", "commerciale")
        self.cmb_type.addItem("Normalizzati (6)", "normalizzato")
        self.cmb_type.setFixedWidth(160)
        self.cmb_type.currentIndexChanged.connect(self._refresh)
        filter_row.addWidget(self.cmb_type)

        filter_row.addWidget(QLabel("Categoria:"))
        self.cmb_cat = QComboBox()
        self.cmb_cat.addItem("Tutte", 0)
        cats = session.commercial.get_categories(only_active=True)
        for cat in cats:
            self.cmb_cat.addItem(
                f"[{cat['code']}] {cat['description']}", cat["id"]
            )
        self.cmb_cat.setFixedWidth(160)
        self.cmb_cat.currentIndexChanged.connect(self._refresh)
        filter_row.addWidget(self.cmb_cat)

        filter_row.addWidget(QLabel("Cerca:"))
        self.txt_search = QLineEdit()
        self.txt_search.setPlaceholderText("Codice, descrizione…")
        self.txt_search.returnPressed.connect(self._refresh)
        filter_row.addWidget(self.txt_search)

        btn_search = QPushButton("↻")
        btn_search.setFixedWidth(28)
        btn_search.clicked.connect(self._refresh)
        filter_row.addWidget(btn_search)
        lay.addLayout(filter_row)

        # Tabella risultati
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(
            ["Codice", "Tipo", "Descrizione", "Fornitore pref."]
        )
        self.table.horizontalHeader().setSectionResizeMode(
            2, QHeaderView.ResizeMode.Stretch
        )
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.doubleClicked.connect(self._accept_selection)
        lay.addWidget(self.table)

        # Bottoni
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("Aggiungi a BOM")
        buttons.accepted.connect(self._accept_selection)
        buttons.rejected.connect(self.reject)
        lay.addWidget(buttons)

    def _refresh(self):
        text      = self.txt_search.text().strip()
        cat_id    = self.cmb_cat.currentData() or 0
        item_type = self.cmb_type.currentData() or ""

        items = session.commercial.search_items(
            text=text,
            category_id=cat_id,
            state="Attivo",
            item_type=item_type,
        )
        self.table.setRowCount(0)
        for it in items:
            r = self.table.rowCount()
            self.table.insertRow(r)
            type_label = "5-COM" if it["item_type"] == "commerciale" else "6-NOR"
            vals = [
                it.get("code") or "",
                type_label,
                it.get("description") or "",
                it.get("preferred_supplier_name") or "",
            ]
            for col, v in enumerate(vals):
                item = QTableWidgetItem(v)
                item.setData(Qt.ItemDataRole.UserRole, it["id"])
                self.table.setItem(r, col, item)

    def _accept_selection(self):
        rows = self.table.selectedItems()
        if not rows:
            return
        self.selected_id = rows[0].data(Qt.ItemDataRole.UserRole)
        self.accept()

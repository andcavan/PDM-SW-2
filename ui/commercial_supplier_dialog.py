# =============================================================================
#  ui/commercial_supplier_dialog.py  –  Gestione fornitori commerciali
# =============================================================================
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
    QGroupBox, QFormLayout, QMessageBox, QCheckBox, QTextEdit,
    QDoubleSpinBox, QSpinBox, QComboBox, QDialogButtonBox,
    QSizePolicy,
)
from PyQt6.QtCore import Qt

from ui.session import session


class CommercialSupplierDialog(QDialog):
    """Dialogo per gestione del registro fornitori/produttori."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Gestione Fornitori / Produttori")
        self.setMinimumSize(820, 520)
        self._editing_id: int | None = None
        self._build_ui()
        self._refresh()

    # ------------------------------------------------------------------
    def _build_ui(self):
        layout = QHBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(16, 16, 16, 16)

        # ── Lista fornitori ────────────────────────────────────────────
        left = QVBoxLayout()

        search_row = QHBoxLayout()
        search_row.addWidget(QLabel("Cerca:"))
        self.txt_search = QLineEdit()
        self.txt_search.setPlaceholderText("Nome, sigla...")
        self.txt_search.textChanged.connect(self._apply_filter)
        search_row.addWidget(self.txt_search)
        left.addLayout(search_row)

        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["Nome", "Sigla", "Email", "Attivo"])
        self.table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeMode.Stretch
        )
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.selectionModel().selectionChanged.connect(self._on_select)
        self.table.doubleClicked.connect(self._on_select)
        left.addWidget(self.table)

        btn_row = QHBoxLayout()
        btn_new = QPushButton("+ Nuovo fornitore")
        btn_new.clicked.connect(self._new_supplier)
        self.btn_deact = QPushButton("Disattiva")
        self.btn_deact.setObjectName("btn_danger")
        self.btn_deact.clicked.connect(self._deactivate)
        btn_row.addWidget(btn_new)
        btn_row.addWidget(self.btn_deact)
        left.addLayout(btn_row)

        layout.addLayout(left, 3)

        # ── Form dettaglio ─────────────────────────────────────────────
        right = QVBoxLayout()
        grp = QGroupBox("Dettaglio fornitore")
        form = QFormLayout(grp)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        self.txt_name    = QLineEdit()
        self.txt_code    = QLineEdit()
        self.txt_code.setMaxLength(10)
        self.txt_contact = QLineEdit()
        self.txt_email   = QLineEdit()
        self.txt_phone   = QLineEdit()
        self.txt_website = QLineEdit()
        self.txt_notes   = QTextEdit()
        self.txt_notes.setMaximumHeight(80)
        self.chk_active  = QCheckBox("Attivo")
        self.chk_active.setChecked(True)

        form.addRow("Nome*:", self.txt_name)
        form.addRow("Sigla:", self.txt_code)
        form.addRow("Referente:", self.txt_contact)
        form.addRow("Email:", self.txt_email)
        form.addRow("Telefono:", self.txt_phone)
        form.addRow("Sito web:", self.txt_website)
        form.addRow("Note:", self.txt_notes)
        form.addRow(self.chk_active)

        right.addWidget(grp)
        right.addStretch()

        self.btn_save = QPushButton("Salva")
        self.btn_save.setObjectName("btn_primary")
        self.btn_save.clicked.connect(self._save)
        right.addWidget(self.btn_save)

        layout.addLayout(right, 2)

    # ------------------------------------------------------------------
    def _refresh(self):
        rows = session.commercial.get_suppliers(only_active=False)
        self._all_rows = rows
        self._populate_table(rows)

    def _populate_table(self, rows: list):
        self.table.setRowCount(0)
        for r in rows:
            row_idx = self.table.rowCount()
            self.table.insertRow(row_idx)
            for col, key in enumerate(["name", "short_code", "email"]):
                item = QTableWidgetItem(r.get(key) or "")
                item.setData(Qt.ItemDataRole.UserRole, r["id"])
                self.table.setItem(row_idx, col, item)
            chk = QTableWidgetItem("✓" if r.get("active") else "✗")
            chk.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(row_idx, 3, chk)

    def _apply_filter(self, text: str):
        text = text.lower()
        for row in range(self.table.rowCount()):
            name  = (self.table.item(row, 0).text() or "").lower()
            sigla = (self.table.item(row, 1).text() or "").lower()
            hidden = text and text not in name and text not in sigla
            self.table.setRowHidden(row, hidden)

    def _on_select(self):
        rows = self.table.selectedItems()
        if not rows:
            return
        sup_id = rows[0].data(Qt.ItemDataRole.UserRole)
        sup = session.commercial.get_supplier(sup_id)
        if not sup:
            return
        self._editing_id = sup_id
        self.txt_name.setText(sup.get("name") or "")
        self.txt_code.setText(sup.get("short_code") or "")
        self.txt_contact.setText(sup.get("contact") or "")
        self.txt_email.setText(sup.get("email") or "")
        self.txt_phone.setText(sup.get("phone") or "")
        self.txt_website.setText(sup.get("website") or "")
        self.txt_notes.setPlainText(sup.get("notes") or "")
        self.chk_active.setChecked(bool(sup.get("active", 1)))

    def _new_supplier(self):
        self._editing_id = None
        self.txt_name.clear()
        self.txt_code.clear()
        self.txt_contact.clear()
        self.txt_email.clear()
        self.txt_phone.clear()
        self.txt_website.clear()
        self.txt_notes.clear()
        self.chk_active.setChecked(True)
        self.txt_name.setFocus()

    def _save(self):
        name = self.txt_name.text().strip()
        if not name:
            QMessageBox.warning(self, "Campo obbligatorio", "Inserire il nome del fornitore.")
            return
        try:
            if self._editing_id is None:
                session.commercial.create_supplier(
                    name=name,
                    short_code=self.txt_code.text().strip(),
                    contact=self.txt_contact.text().strip(),
                    email=self.txt_email.text().strip(),
                    phone=self.txt_phone.text().strip(),
                    website=self.txt_website.text().strip(),
                    notes=self.txt_notes.toPlainText().strip(),
                )
            else:
                session.commercial.update_supplier(
                    sup_id=self._editing_id,
                    name=name,
                    short_code=self.txt_code.text().strip(),
                    contact=self.txt_contact.text().strip(),
                    email=self.txt_email.text().strip(),
                    phone=self.txt_phone.text().strip(),
                    website=self.txt_website.text().strip(),
                    notes=self.txt_notes.toPlainText().strip(),
                    active=1 if self.chk_active.isChecked() else 0,
                )
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))
            return
        self._refresh()
        self._new_supplier()

    def _deactivate(self):
        rows = self.table.selectedItems()
        if not rows:
            return
        sup_id = rows[0].data(Qt.ItemDataRole.UserRole)
        sup = session.commercial.get_supplier(sup_id)
        if not sup:
            return
        reply = QMessageBox.question(
            self, "Conferma",
            f"Disattivare il fornitore '{sup['name']}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            session.commercial.deactivate_supplier(sup_id)
            self._refresh()


# =============================================================================
#  CommercialSupplierLinkDialog  –  Mini-dialog per collegare fornitore ad articolo
# =============================================================================

class CommercialSupplierLinkDialog(QDialog):
    """Dialogo per aggiungere/modificare un collegamento articolo-fornitore."""

    def __init__(self, item_id: int, link_data: dict | None = None, parent=None):
        """
        Args:
            item_id:   ID articolo commerciale
            link_data: dict esistente (edit mode) o None (create mode)
        """
        super().__init__(parent)
        self.item_id   = item_id
        self.link_data = link_data  # None = nuova associazione
        self.setWindowTitle(
            "Modifica fornitore" if link_data else "Aggiungi fornitore"
        )
        self.setMinimumWidth(420)
        self._build_ui()
        if link_data:
            self._load(link_data)

    # ------------------------------------------------------------------
    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        # Fornitore (combo solo in creazione, label in modifica)
        self.cmb_supplier = QComboBox()
        self._suppliers = session.commercial.get_suppliers(only_active=True)
        for s in self._suppliers:
            label = s["name"]
            if s.get("short_code"):
                label += f" [{s['short_code']}]"
            self.cmb_supplier.addItem(label, s["id"])

        self.lbl_supplier = QLabel()
        if self.link_data:
            self.cmb_supplier.hide()
            form.addRow("Fornitore:", self.lbl_supplier)
        else:
            self.lbl_supplier.hide()
            form.addRow("Fornitore*:", self.cmb_supplier)

        self.txt_supplier_code = QLineEdit()
        self.txt_supplier_code.setPlaceholderText("Codice articolo del fornitore")

        self.spin_price = QDoubleSpinBox()
        self.spin_price.setRange(0.0, 999999.99)
        self.spin_price.setDecimals(4)
        self.spin_price.setSpecialValueText("—")

        self.cmb_currency = QComboBox()
        self.cmb_currency.addItems(["EUR", "USD", "GBP", "CHF"])

        price_row = QHBoxLayout()
        price_row.addWidget(self.spin_price)
        price_row.addWidget(self.cmb_currency)

        self.spin_lead = QSpinBox()
        self.spin_lead.setRange(0, 999)
        self.spin_lead.setSuffix(" giorni")
        self.spin_lead.setSpecialValueText("—")

        self.chk_preferred = QCheckBox("Fornitore preferenziale")

        self.txt_notes = QLineEdit()
        self.txt_notes.setPlaceholderText("Note opzionali")

        form.addRow("Codice fornitore:", self.txt_supplier_code)
        form.addRow("Prezzo unitario:", price_row)
        form.addRow("Lead time:", self.spin_lead)
        form.addRow(self.chk_preferred)
        form.addRow("Note:", self.txt_notes)

        layout.addLayout(form)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self._save)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _load(self, d: dict):
        self.lbl_supplier.setText(d.get("supplier_name") or "")
        self.txt_supplier_code.setText(d.get("supplier_code") or "")
        price = d.get("unit_price") or 0.0
        self.spin_price.setValue(price)
        idx = self.cmb_currency.findText(d.get("currency") or "EUR")
        if idx >= 0:
            self.cmb_currency.setCurrentIndex(idx)
        self.spin_lead.setValue(d.get("lead_time_days") or 0)
        self.chk_preferred.setChecked(bool(d.get("is_preferred")))
        self.txt_notes.setText(d.get("notes") or "")

    def _save(self):
        try:
            price = self.spin_price.value() if self.spin_price.value() > 0 else None
            lead  = self.spin_lead.value()  if self.spin_lead.value()  > 0 else None

            if self.link_data:
                session.commercial.update_item_supplier(
                    link_id=self.link_data["id"],
                    supplier_code=self.txt_supplier_code.text().strip(),
                    unit_price=price,
                    currency=self.cmb_currency.currentText(),
                    lead_time_days=lead,
                    is_preferred=self.chk_preferred.isChecked(),
                    notes=self.txt_notes.text().strip(),
                )
            else:
                sup_id = self.cmb_supplier.currentData()
                if sup_id is None:
                    QMessageBox.warning(self, "Campo obbligatorio",
                                        "Selezionare un fornitore.")
                    return
                # Verifica che il fornitore non sia già collegato
                existing = [
                    s["supplier_id"]
                    for s in session.commercial.get_item_suppliers(self.item_id)
                ]
                if sup_id in existing:
                    QMessageBox.warning(
                        self, "Duplicato",
                        "Questo fornitore è già collegato all'articolo."
                    )
                    return
                session.commercial.add_item_supplier(
                    item_id=self.item_id,
                    supplier_id=sup_id,
                    supplier_code=self.txt_supplier_code.text().strip(),
                    unit_price=price,
                    currency=self.cmb_currency.currentText(),
                    lead_time_days=lead,
                    is_preferred=self.chk_preferred.isChecked(),
                    notes=self.txt_notes.text().strip(),
                )
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))
            return
        self.accept()

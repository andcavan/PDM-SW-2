# =============================================================================
#  ui/commercial_item_dialog.py  –  Creazione/modifica articolo commerciale
# =============================================================================
from pathlib import Path

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
    QGroupBox, QFormLayout, QMessageBox, QTextEdit, QTabWidget,
    QDialogButtonBox, QComboBox, QWidget, QFileDialog, QSizePolicy,
    QDoubleSpinBox,
)
from PyQt6.QtCore import Qt, pyqtSignal

from config import COMMERCIAL_ITEM_TYPES, COMMERCIAL_WORKFLOW_TRANSITIONS
from ui.session import session


_ROLE_LINK_ID = Qt.ItemDataRole.UserRole
_ROLE_SUP_ID  = Qt.ItemDataRole.UserRole + 1


class CommercialItemDialog(QDialog):
    """Dialogo per creare o modificare un articolo commerciale/normalizzato."""

    saved = pyqtSignal(int)  # emesso con item_id dopo salvataggio

    def __init__(self, item_id: int | None = None, parent=None,
                 prefill: dict | None = None):
        super().__init__(parent)
        self.item_id    = item_id
        self._prefill   = prefill        # dati pre-compilati per duplicazione
        self._sw_file_pending: Path | None = None
        self._is_edit   = item_id is not None
        self.setWindowTitle(
            "Modifica articolo" if self._is_edit
            else ("Duplica articolo" if prefill else "Nuovo articolo commerciale")
        )
        self.setMinimumSize(680, 560)
        self._build_ui()
        if self._is_edit:
            self._load_item()
        elif prefill:
            self._apply_prefill(prefill)
        else:
            self._update_code_preview()

    # ==================================================================
    #  Costruzione UI
    # ==================================================================

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setSpacing(10)
        root.setContentsMargins(14, 14, 14, 14)

        self.tabs = QTabWidget()
        self.tabs.addTab(self._tab_generale(),   "Generale")
        self.tabs.addTab(self._tab_fornitori(),  "Fornitori")
        self.tabs.addTab(self._tab_sw_props(),   "Proprietà SW")
        self.tabs.addTab(self._tab_storico(),    "Storico")
        root.addWidget(self.tabs)

        btn_row = QHBoxLayout()

        # Bottone Duplica (solo in modifica)
        if self._is_edit:
            self.btn_duplicate = QPushButton("Duplica articolo")
            self.btn_duplicate.setToolTip(
                "Crea una copia di questo articolo con nuovo codice"
            )
            self.btn_duplicate.clicked.connect(self._action_duplicate)
            btn_row.addWidget(self.btn_duplicate)

        btn_row.addStretch()

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save
            | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.button(QDialogButtonBox.StandardButton.Save).setText("Salva")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Annulla")
        buttons.accepted.connect(self._save)
        buttons.rejected.connect(self.reject)
        btn_row.addWidget(buttons)

        root.addLayout(btn_row)

    # ── Tab Generale ──────────────────────────────────────────────────
    def _tab_generale(self) -> QWidget:
        w = QWidget()
        form = QFormLayout(w)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        form.setSpacing(10)
        form.setContentsMargins(14, 14, 14, 14)

        # Tipo
        self.cmb_type = QComboBox()
        self.cmb_type.addItem("Commerciale (5)", "commerciale")
        self.cmb_type.addItem("Normalizzato (6)", "normalizzato")
        self.cmb_type.currentIndexChanged.connect(self._on_type_changed)
        form.addRow("Tipo*:", self.cmb_type)

        # Categoria
        self.cmb_cat = QComboBox()
        self._populate_categories()
        self.cmb_cat.currentIndexChanged.connect(self._on_cat_changed)
        form.addRow("Categoria*:", self.cmb_cat)

        # Sottocategoria (obbligatoria)
        self.cmb_sub = QComboBox()
        self.cmb_sub.addItem("— Selezionare —", None)
        self.cmb_sub.currentIndexChanged.connect(self._on_sub_changed)
        form.addRow("Sottocategoria*:", self.cmb_sub)

        # Codice (preview)
        code_row = QHBoxLayout()
        self.lbl_code = QLabel("—")
        self.lbl_code.setObjectName("code_preview")
        code_row.addWidget(self.lbl_code)
        code_row.addStretch()
        form.addRow("Codice:", code_row)

        # Descrizione + template
        self.txt_desc = QLineEdit()
        self.txt_desc.setPlaceholderText("Descrizione articolo")
        desc_row = QHBoxLayout()
        desc_row.addWidget(self.txt_desc)
        self.btn_use_tpl = QPushButton("Usa template")
        self.btn_use_tpl.setMaximumWidth(110)
        self.btn_use_tpl.clicked.connect(self._apply_template)
        self.btn_use_tpl.setEnabled(False)
        desc_row.addWidget(self.btn_use_tpl)
        form.addRow("Descrizione*:", desc_row)

        self.lbl_template = QLabel()
        self.lbl_template.setWordWrap(True)
        self.lbl_template.hide()
        form.addRow("", self.lbl_template)

        # Note
        self.txt_notes = QTextEdit()
        self.txt_notes.setMaximumHeight(70)
        self.txt_notes.setPlaceholderText("Note opzionali")
        form.addRow("Note:", self.txt_notes)

        # Stato (solo edit)
        self.lbl_state = QLabel("Attivo")
        if self._is_edit:
            form.addRow("Stato:", self.lbl_state)

        # Collegamento file SW
        grp_sw = QGroupBox("File SolidWorks (opzionale)")
        sw_lay = QHBoxLayout(grp_sw)
        self.lbl_sw_file = QLabel("Nessun file collegato")
        self.lbl_sw_file.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )
        sw_lay.addWidget(self.lbl_sw_file)
        self.btn_sw_select = QPushButton("Seleziona…")
        self.btn_sw_remove = QPushButton("Rimuovi")
        self.btn_sw_remove.setEnabled(False)
        self.btn_sw_select.clicked.connect(self._select_sw_file)
        self.btn_sw_remove.clicked.connect(self._remove_sw_file)
        sw_lay.addWidget(self.btn_sw_select)
        sw_lay.addWidget(self.btn_sw_remove)
        form.addRow(grp_sw)

        return w

    # ── Tab Fornitori ─────────────────────────────────────────────────
    def _tab_fornitori(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(10, 10, 10, 10)

        self.tbl_suppliers = QTableWidget(0, 6)
        self.tbl_suppliers.setHorizontalHeaderLabels(
            ["Fornitore", "Cod. fornitore", "Prezzo", "Val.", "Lead (gg)", "Pref."]
        )
        self.tbl_suppliers.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeMode.Stretch
        )
        self.tbl_suppliers.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows
        )
        self.tbl_suppliers.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tbl_suppliers.doubleClicked.connect(self._edit_supplier_link)
        lay.addWidget(self.tbl_suppliers)

        btn_row = QHBoxLayout()
        btn_add  = QPushButton("+ Aggiungi fornitore")
        btn_edit = QPushButton("Modifica")
        btn_rem  = QPushButton("Rimuovi")
        btn_pref = QPushButton("★ Imposta preferito")
        btn_add.clicked.connect(self._add_supplier_link)
        btn_edit.clicked.connect(self._edit_supplier_link)
        btn_rem.clicked.connect(self._remove_supplier_link)
        btn_pref.clicked.connect(self._set_preferred)
        for btn in (btn_add, btn_edit, btn_rem, btn_pref):
            btn_row.addWidget(btn)
        lay.addLayout(btn_row)

        return w

    # ── Tab Proprietà SW ──────────────────────────────────────────────
    def _tab_sw_props(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(10, 10, 10, 10)

        lay.addWidget(QLabel(
            "Sincronizzazione proprietà SolidWorks. "
            "Richiede file SW collegato e SolidWorks aperto."
        ))

        ctrl_row = QHBoxLayout()
        self.btn_import_sw = QPushButton("← Importa da SW")
        self.btn_export_sw = QPushButton("→ Esporta verso SW")
        self.btn_bidir_sw  = QPushButton("↔ Bidirezionale")
        self.btn_import_sw.clicked.connect(self._import_from_sw)
        self.btn_export_sw.clicked.connect(self._export_to_sw)
        self.btn_bidir_sw.clicked.connect(self._bidir_sw)
        for btn in (self.btn_import_sw, self.btn_export_sw, self.btn_bidir_sw):
            btn.setEnabled(False)
            ctrl_row.addWidget(btn)
        ctrl_row.addStretch()
        lay.addLayout(ctrl_row)

        self.tbl_props = QTableWidget(0, 2)
        self.tbl_props.setHorizontalHeaderLabels(["Proprietà", "Valore"])
        self.tbl_props.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeMode.Stretch
        )
        self.tbl_props.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        lay.addWidget(self.tbl_props)

        return w

    # ── Tab Storico ───────────────────────────────────────────────────
    def _tab_storico(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(10, 10, 10, 10)

        self.tbl_log = QTableWidget(0, 4)
        self.tbl_log.setHorizontalHeaderLabels(
            ["Data", "Utente", "Azione", "Note"]
        )
        self.tbl_log.horizontalHeader().setSectionResizeMode(
            3, QHeaderView.ResizeMode.Stretch
        )
        self.tbl_log.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        lay.addWidget(self.tbl_log)

        return w

    # ==================================================================
    #  Caricamento dati
    # ==================================================================

    def _populate_categories(self):
        item_type = self.cmb_type.currentData() or ""
        self.cmb_cat.blockSignals(True)
        self.cmb_cat.clear()
        cats = session.commercial.get_categories(
            item_type=item_type if item_type else None
        )
        for cat in cats:
            self.cmb_cat.addItem(
                f"[{cat['code']}] {cat['description']}", cat["id"]
            )
        self.cmb_cat.blockSignals(False)

    def _populate_subcategories(self, cat_id: int | None):
        self.cmb_sub.blockSignals(True)
        self.cmb_sub.clear()
        self.cmb_sub.addItem("— Selezionare —", None)
        if cat_id:
            subs = session.commercial.get_subcategories(cat_id)
            for sub in subs:
                self.cmb_sub.addItem(
                    f"[{sub['code']}] {sub['description']}", sub["id"]
                )
        self.cmb_sub.blockSignals(False)
        self._update_code_preview()
        self._update_template_hint()

    def _on_type_changed(self):
        """Ricarica le categorie filtrate per il tipo selezionato."""
        self._populate_categories()
        self._on_cat_changed()

    def _on_cat_changed(self):
        cat_id = self.cmb_cat.currentData()
        self._populate_subcategories(cat_id)

    def _on_sub_changed(self):
        self._update_code_preview()
        self._update_template_hint()

    def _update_code_preview(self):
        if self._is_edit:
            return  # codice fisso in modifica
        cat_id = self.cmb_cat.currentData()
        sub_id = self.cmb_sub.currentData()
        item_type = self.cmb_type.currentData() or "commerciale"
        if cat_id:
            code = session.commercial.preview_code(item_type, cat_id, sub_id)
        else:
            code = "—"
        self.lbl_code.setText(code)

    def _update_template_hint(self):
        sub_id = self.cmb_sub.currentData()
        if sub_id:
            sub = session.commercial.get_subcategory(sub_id)
            tpl = sub.get("desc_template") if sub else None
            if tpl:
                self.lbl_template.setText(f"Template: {tpl}")
                self.lbl_template.show()
                self.btn_use_tpl.setEnabled(True)
                return
        self.lbl_template.hide()
        self.btn_use_tpl.setEnabled(False)

    def _apply_template(self):
        sub_id = self.cmb_sub.currentData()
        if not sub_id:
            return
        sub = session.commercial.get_subcategory(sub_id)
        if sub and sub.get("desc_template"):
            self.txt_desc.setText(sub["desc_template"])
            self.txt_desc.setFocus()

    def _load_item(self):
        """Carica dati di un articolo esistente."""
        item = session.commercial.get_item(self.item_id)
        if not item:
            return

        # Tipo
        idx = self.cmb_type.findData(item.get("item_type", "commerciale"))
        if idx >= 0:
            self.cmb_type.setCurrentIndex(idx)
        self.cmb_type.setEnabled(False)  # non modificabile dopo creazione

        # Categoria
        cat_idx = self.cmb_cat.findData(item.get("category_id"))
        if cat_idx >= 0:
            self.cmb_cat.setCurrentIndex(cat_idx)
        self.cmb_cat.setEnabled(False)

        # Sottocategoria
        self._populate_subcategories(item.get("category_id"))
        sub_idx = self.cmb_sub.findData(item.get("subcategory_id"))
        if sub_idx >= 0:
            self.cmb_sub.setCurrentIndex(sub_idx)
        self.cmb_sub.setEnabled(False)

        # Codice (fisso)
        self.lbl_code.setText(item.get("code") or "—")

        # Campi testo
        self.txt_desc.setText(item.get("description") or "")
        self.txt_notes.setPlainText(item.get("notes") or "")
        self.lbl_state.setText(item.get("state") or "")

        # File SW
        if item.get("file_name"):
            self.lbl_sw_file.setText(item["file_name"])
            self.btn_sw_remove.setEnabled(True)
            for btn in (self.btn_import_sw, self.btn_export_sw, self.btn_bidir_sw):
                btn.setEnabled(True)

        # Fornitori
        self._load_suppliers()

        # Proprietà SW
        self._load_props()

        # Storico
        self._load_log()

    def _apply_prefill(self, src: dict):
        """Pre-compila i campi con i dati dell'articolo sorgente (per duplicazione)."""
        # Tipo
        idx = self.cmb_type.findData(src.get("item_type", "commerciale"))
        if idx >= 0:
            self.cmb_type.setCurrentIndex(idx)

        # Categoria
        cat_idx = self.cmb_cat.findData(src.get("category_id"))
        if cat_idx >= 0:
            self.cmb_cat.setCurrentIndex(cat_idx)

        # Sottocategoria
        self._populate_subcategories(src.get("category_id"))
        sub_idx = self.cmb_sub.findData(src.get("subcategory_id"))
        if sub_idx >= 0:
            self.cmb_sub.setCurrentIndex(sub_idx)

        # Testo
        self.txt_desc.setText(src.get("description") or "")
        self.txt_notes.setPlainText(src.get("notes") or "")

        self._update_code_preview()
        self._update_template_hint()

    def _load_suppliers(self):
        self.tbl_suppliers.setRowCount(0)
        if not self.item_id:
            return
        links = session.commercial.get_item_suppliers(self.item_id)
        for lnk in links:
            r = self.tbl_suppliers.rowCount()
            self.tbl_suppliers.insertRow(r)
            price_str = (
                f"{lnk['unit_price']:.4f}" if lnk.get("unit_price") is not None else "—"
            )
            lead_str  = str(lnk["lead_time_days"]) if lnk.get("lead_time_days") else "—"
            pref_str  = "★" if lnk.get("is_preferred") else ""
            vals = [
                lnk.get("supplier_name") or "",
                lnk.get("supplier_code") or "",
                price_str,
                lnk.get("currency") or "EUR",
                lead_str,
                pref_str,
            ]
            for col, v in enumerate(vals):
                it = QTableWidgetItem(v)
                it.setData(_ROLE_LINK_ID, lnk["id"])
                it.setData(_ROLE_SUP_ID,  lnk["supplier_id"])
                self.tbl_suppliers.setItem(r, col, it)

    def _load_props(self):
        self.tbl_props.setRowCount(0)
        if not self.item_id:
            return
        props = session.commercial.get_properties(self.item_id)
        for name, value in sorted(props.items()):
            r = self.tbl_props.rowCount()
            self.tbl_props.insertRow(r)
            self.tbl_props.setItem(r, 0, QTableWidgetItem(name))
            self.tbl_props.setItem(r, 1, QTableWidgetItem(value or ""))

    def _load_log(self):
        self.tbl_log.setRowCount(0)
        if not self.item_id:
            return
        logs = session.commercial.get_checkout_log(self.item_id)
        for entry in logs:
            r = self.tbl_log.rowCount()
            self.tbl_log.insertRow(r)
            ts = (entry.get("timestamp") or "")[:19].replace("T", " ")
            self.tbl_log.setItem(r, 0, QTableWidgetItem(ts))
            self.tbl_log.setItem(r, 1, QTableWidgetItem(entry.get("user_name") or ""))
            self.tbl_log.setItem(r, 2, QTableWidgetItem(entry.get("action") or ""))
            self.tbl_log.setItem(r, 3, QTableWidgetItem(entry.get("notes") or ""))

    # ==================================================================
    #  Gestione file SolidWorks
    # ==================================================================

    def _select_sw_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Seleziona file SolidWorks",
            "", "SolidWorks (*.SLDPRT *.SLDASM *.sldprt *.sldasm)"
        )
        if not path:
            return
        self._sw_file_pending = Path(path)
        self.lbl_sw_file.setText(self._sw_file_pending.name)
        self.btn_sw_remove.setEnabled(True)
        for btn in (self.btn_import_sw, self.btn_export_sw, self.btn_bidir_sw):
            btn.setEnabled(True)

    def _remove_sw_file(self):
        self._sw_file_pending = None
        self.lbl_sw_file.setText("Nessun file collegato")
        self.btn_sw_remove.setEnabled(False)
        for btn in (self.btn_import_sw, self.btn_export_sw, self.btn_bidir_sw):
            btn.setEnabled(False)
        if self._is_edit and self.item_id:
            reply = QMessageBox.question(
                self, "Conferma",
                "Rimuovere il collegamento al file SolidWorks?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if reply == QMessageBox.StandardButton.Yes:
                session.commercial.unlink_sw_file(
                    self.item_id, session.user
                )

    # ==================================================================
    #  Gestione fornitori
    # ==================================================================

    def _add_supplier_link(self):
        if not self.item_id:
            QMessageBox.information(
                self, "Salva prima",
                "Salvare l'articolo prima di aggiungere fornitori."
            )
            return
        from ui.commercial_supplier_dialog import CommercialSupplierLinkDialog
        dlg = CommercialSupplierLinkDialog(self.item_id, parent=self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self._load_suppliers()

    def _edit_supplier_link(self):
        row = self.tbl_suppliers.currentRow()
        if row < 0:
            return
        link_id = self.tbl_suppliers.item(row, 0).data(_ROLE_LINK_ID)
        links = session.commercial.get_item_suppliers(self.item_id)
        link_data = next((l for l in links if l["id"] == link_id), None)
        if not link_data:
            return
        from ui.commercial_supplier_dialog import CommercialSupplierLinkDialog
        dlg = CommercialSupplierLinkDialog(self.item_id, link_data=link_data, parent=self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self._load_suppliers()

    def _remove_supplier_link(self):
        row = self.tbl_suppliers.currentRow()
        if row < 0:
            return
        link_id = self.tbl_suppliers.item(row, 0).data(_ROLE_LINK_ID)
        reply = QMessageBox.question(
            self, "Conferma",
            "Rimuovere il fornitore dall'articolo?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            session.commercial.remove_item_supplier(link_id)
            self._load_suppliers()

    def _set_preferred(self):
        row = self.tbl_suppliers.currentRow()
        if row < 0:
            return
        sup_id = self.tbl_suppliers.item(row, 0).data(_ROLE_SUP_ID)
        session.commercial.set_preferred_supplier(self.item_id, sup_id)
        self._load_suppliers()

    # ==================================================================
    #  Sincronizzazione SW
    # ==================================================================

    def _get_sw_path(self) -> Path | None:
        if self._sw_file_pending:
            return self._sw_file_pending
        if self.item_id:
            item = session.commercial.get_item(self.item_id)
            if item and item.get("archive_path") and session.sp:
                p = session.sp.root / item["archive_path"]
                return p if p.exists() else None
        return None

    def _import_from_sw(self):
        sw_path = self._get_sw_path()
        if not sw_path:
            QMessageBox.warning(self, "File mancante", "Nessun file SolidWorks trovato.")
            return
        result = session.commercial.sync_sw_to_pdm(
            self.item_id, sw_path, session.properties
        )
        if result.get("ok"):
            self._load_props()
            QMessageBox.information(
                self, "Importazione completata",
                f"{result['imported_count']} proprietà importate da SolidWorks."
            )
        else:
            QMessageBox.critical(self, "Errore", result.get("error", "Errore sconosciuto"))

    def _export_to_sw(self):
        sw_path = self._get_sw_path()
        if not sw_path:
            QMessageBox.warning(self, "File mancante", "Nessun file SolidWorks trovato.")
            return
        result = session.commercial.sync_pdm_to_sw(
            self.item_id, sw_path, session.properties
        )
        if result.get("ok"):
            QMessageBox.information(
                self, "Esportazione completata",
                f"{result['written_count']} proprietà scritte in SolidWorks."
            )
        else:
            QMessageBox.critical(self, "Errore", result.get("error", "Errore sconosciuto"))

    def _bidir_sw(self):
        self._import_from_sw()
        self._export_to_sw()

    # ==================================================================
    #  Duplicazione
    # ==================================================================

    def _action_duplicate(self):
        """Apre un dialog pre-compilato per creare una copia dell'articolo corrente."""
        if not self.item_id:
            return
        item = session.commercial.get_item(self.item_id)
        if not item:
            return
        dlg = CommercialItemDialog(prefill=dict(item), parent=self.parent())
        dlg.saved.connect(self.saved)
        if dlg.exec():
            self.saved.emit(self.item_id)

    # ==================================================================
    #  Salvataggio
    # ==================================================================

    def _save(self):
        desc = self.txt_desc.text().strip()
        if not desc:
            QMessageBox.warning(self, "Campo obbligatorio",
                                "Inserire la descrizione dell'articolo.")
            return

        cat_id    = self.cmb_cat.currentData()
        sub_id    = self.cmb_sub.currentData()
        item_type = self.cmb_type.currentData() or "commerciale"
        notes     = self.txt_notes.toPlainText().strip()
        user_id   = session.user["id"] if session.user else None

        if not cat_id:
            QMessageBox.warning(self, "Campo obbligatorio",
                                "Selezionare una categoria.")
            return

        if not sub_id:
            QMessageBox.warning(self, "Campo obbligatorio",
                                "Selezionare una sottocategoria.")
            return

        try:
            if self._is_edit:
                session.commercial.update_item(
                    self.item_id, desc, notes, modified_by=user_id
                )
            else:
                self.item_id = session.commercial.create_item(
                    item_type=item_type,
                    category_id=cat_id,
                    subcategory_id=sub_id,
                    description=desc,
                    notes=notes,
                    created_by=user_id,
                )
                self._is_edit = True

            # Collega file SW se selezionato
            if self._sw_file_pending and session.sp:
                session.commercial.link_sw_file(
                    self.item_id,
                    self._sw_file_pending,
                    session.sp,
                    current_user=session.user,
                )
                self._sw_file_pending = None

        except Exception as e:
            QMessageBox.critical(self, "Errore salvataggio", str(e))
            return

        self.saved.emit(self.item_id)
        self.accept()

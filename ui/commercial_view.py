# =============================================================================
#  ui/commercial_view.py  –  Vista principale articoli commerciali/normalizzati
# =============================================================================
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QTreeWidget, QTreeWidgetItem, QHeaderView,
    QComboBox, QSplitter, QMenu, QMessageBox, QGroupBox,
    QTableWidget, QTableWidgetItem, QDialog, QTextEdit,
    QFormLayout, QFrame, QScrollArea,
)
from PyQt6.QtCore import Qt, pyqtSignal, QSettings
from PyQt6.QtGui import QAction, QFont

from config import COMMERCIAL_ITEM_TYPES, COMMERCIAL_WORKFLOW_TRANSITIONS
from ui.session import session

_ROLE_ID   = Qt.ItemDataRole.UserRole
_ROLE_TYPE = Qt.ItemDataRole.UserRole + 1  # 'cat' | 'sub' | 'item'

_APP = "PDM-SW"
_KEY = "CommercialView"


class CommercialView(QWidget):
    """Tab principale gestione articoli commerciali e normalizzati."""

    item_selected = pyqtSignal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._build_ui()
        self._populate_filters()

    # ==================================================================
    #  Costruzione UI
    # ==================================================================

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setSpacing(6)
        root.setContentsMargins(10, 10, 10, 10)

        # ── Barra filtri ───────────────────────────────────────────────
        filter_bar = QHBoxLayout()
        filter_bar.setSpacing(8)

        filter_bar.addWidget(QLabel("Tipo:"))
        self.cmb_type = QComboBox()
        self.cmb_type.setFixedWidth(150)
        self.cmb_type.addItem("Tutti i tipi", "")
        self.cmb_type.addItem("Commerciali (5)", "commerciale")
        self.cmb_type.addItem("Normalizzati (6)", "normalizzato")
        self.cmb_type.currentIndexChanged.connect(self._on_type_filter_changed)
        filter_bar.addWidget(self.cmb_type)

        filter_bar.addWidget(QLabel("Cerca:"))
        self.txt_search = QLineEdit()
        self.txt_search.setPlaceholderText("Codice, descrizione…")
        self.txt_search.setFixedWidth(180)
        self.txt_search.returnPressed.connect(self.refresh)
        filter_bar.addWidget(self.txt_search)

        filter_bar.addWidget(QLabel("Categoria:"))
        self.cmb_cat = QComboBox()
        self.cmb_cat.setFixedWidth(160)
        self.cmb_cat.currentIndexChanged.connect(self._on_cat_filter_changed)
        filter_bar.addWidget(self.cmb_cat)

        filter_bar.addWidget(QLabel("Sottocategoria:"))
        self.cmb_sub = QComboBox()
        self.cmb_sub.setFixedWidth(160)
        self.cmb_sub.currentIndexChanged.connect(self.refresh)
        filter_bar.addWidget(self.cmb_sub)

        filter_bar.addWidget(QLabel("Fornitore:"))
        self.cmb_supplier = QComboBox()
        self.cmb_supplier.setFixedWidth(160)
        self.cmb_supplier.currentIndexChanged.connect(self.refresh)
        filter_bar.addWidget(self.cmb_supplier)

        filter_bar.addWidget(QLabel("Stato:"))
        self.cmb_state = QComboBox()
        self.cmb_state.addItems(["Tutti", "Attivo", "Obsoleto"])
        self.cmb_state.setFixedWidth(90)
        self.cmb_state.currentIndexChanged.connect(self.refresh)
        filter_bar.addWidget(self.cmb_state)

        filter_bar.addStretch()

        root.addLayout(filter_bar)

        # ── Splitter albero / dettaglio ────────────────────────────────
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setObjectName("comm_splitter")

        # Albero
        left = QWidget()
        left_lay = QVBoxLayout(left)
        left_lay.setContentsMargins(0, 0, 0, 0)

        self.tree = QTreeWidget()
        self.tree.setColumnCount(6)
        self.tree.setHeaderLabels(
            ["Codice / Categoria / Sottogruppo", "Descrizione", "Tipo", "Stato", "Fornitore pref.", "🔒"]
        )
        self.tree.header().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.tree.header().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.tree.header().setStretchLastSection(False)
        for col in (2, 3, 4, 5):
            self.tree.header().setSectionResizeMode(
                col, QHeaderView.ResizeMode.ResizeToContents
            )
        self.tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self._context_menu)
        self.tree.currentItemChanged.connect(self._on_tree_item_changed)
        self.tree.itemDoubleClicked.connect(self._on_double_click)
        left_lay.addWidget(self.tree)

        self.lbl_count = QLabel("0 articoli")
        left_lay.addWidget(self.lbl_count)

        splitter.addWidget(left)

        # Pannello dettaglio
        self._detail = CommercialDetailPanel()
        splitter.addWidget(self._detail)
        splitter.setSizes([600, 320])

        # Ripristina posizione splitter
        settings = QSettings(_APP, _KEY)
        state = settings.value("splitter_state")
        if state:
            splitter.restoreState(state)

        self._splitter = splitter
        root.addWidget(splitter)

    # ==================================================================
    #  Filtri
    # ==================================================================

    def _populate_filters(self):
        item_type = self.cmb_type.currentData() or ""
        # Categoria — filtra per tipo selezionato
        self.cmb_cat.blockSignals(True)
        self.cmb_cat.clear()
        self.cmb_cat.addItem("Tutte", 0)
        cats = session.commercial.get_categories(
            item_type=item_type if item_type else None
        )
        for cat in cats:
            self.cmb_cat.addItem(
                f"[{cat['code']}] {cat['description']}", cat["id"]
            )
        self.cmb_cat.blockSignals(False)

        # Sottocategoria
        self._populate_subcategory_filter(0)

        # Fornitore
        self.cmb_supplier.blockSignals(True)
        self.cmb_supplier.clear()
        self.cmb_supplier.addItem("Tutti", 0)
        sups = session.commercial.get_suppliers(only_active=True)
        for s in sups:
            self.cmb_supplier.addItem(s["name"], s["id"])
        self.cmb_supplier.blockSignals(False)

    def _on_type_filter_changed(self):
        """Aggiorna i filtri di categoria al cambio tipo, poi fa il refresh."""
        self._populate_filters()
        self.refresh()

    def _on_cat_filter_changed(self):
        cat_id = self.cmb_cat.currentData() or 0
        self._populate_subcategory_filter(cat_id)
        self.refresh()

    def _populate_subcategory_filter(self, cat_id: int):
        self.cmb_sub.blockSignals(True)
        self.cmb_sub.clear()
        self.cmb_sub.addItem("Tutte", 0)
        if cat_id:
            subs = session.commercial.get_subcategories(cat_id, only_active=True)
            for sub in subs:
                self.cmb_sub.addItem(
                    f"[{sub['code']}] {sub['description']}", sub["id"]
                )
        self.cmb_sub.blockSignals(False)

    # ==================================================================
    #  Refresh
    # ==================================================================

    def refresh(self):
        if not session.commercial:
            return

        text      = self.txt_search.text().strip()
        cat_id    = self.cmb_cat.currentData() or 0
        sub_id    = self.cmb_sub.currentData() or 0
        sup_id    = self.cmb_supplier.currentData() or 0
        state     = self.cmb_state.currentText()
        item_type = self.cmb_type.currentData() or ""

        state_filter = "" if state == "Tutti" else state

        items = session.commercial.search_items(
            text=text,
            category_id=cat_id,
            subcategory_id=sub_id,
            supplier_id=sup_id,
            state=state_filter,
            item_type=item_type,
        )

        self.tree.clear()
        # Raggruppa per categoria → sottogruppo → articolo
        cat_nodes: dict[int, QTreeWidgetItem] = {}
        sub_nodes: dict[tuple, QTreeWidgetItem] = {}  # (cat_id, sub_id)

        for it in items:
            cid = it["category_id"]
            sid = it.get("subcategory_id")

            # Nodo categoria
            if cid not in cat_nodes:
                cat_item = QTreeWidgetItem([
                    f"[{it.get('cat_code','')}] {it.get('cat_description','')}",
                    "", "", "", "", "",
                ])
                cat_item.setData(0, _ROLE_ID,   cid)
                cat_item.setData(0, _ROLE_TYPE, "cat")
                font = cat_item.font(0)
                font.setBold(True)
                cat_item.setFont(0, font)
                self.tree.addTopLevelItem(cat_item)
                cat_nodes[cid] = cat_item

            # Nodo sottogruppo (se presente)
            parent_node = cat_nodes[cid]
            if sid:
                key = (cid, sid)
                if key not in sub_nodes:
                    sub_item = QTreeWidgetItem([
                        f"[{it.get('sub_code','')}] {it.get('sub_description','')}",
                        "", "", "", "", "",
                    ])
                    sub_item.setData(0, _ROLE_ID,   sid)
                    sub_item.setData(0, _ROLE_TYPE, "sub")
                    font = sub_item.font(0)
                    font.setBold(True)
                    sub_item.setFont(0, font)
                    cat_nodes[cid].addChild(sub_item)
                    sub_nodes[key] = sub_item
                parent_node = sub_nodes[key]

            lock_icon = "🔒" if it.get("is_locked") else ""
            type_label = "5-COM" if it["item_type"] == "commerciale" else "6-NOR"
            node = QTreeWidgetItem([
                it["code"],
                it.get("description") or "",
                type_label,
                it.get("state") or "",
                it.get("preferred_supplier_name") or "",
                lock_icon,
            ])
            node.setData(0, _ROLE_ID,   it["id"])
            node.setData(0, _ROLE_TYPE, "item")
            parent_node.addChild(node)

        self.tree.expandAll()

        # Aggiorna conteggio
        n_com = sum(1 for i in items if i["item_type"] == "commerciale")
        n_nor = sum(1 for i in items if i["item_type"] == "normalizzato")
        self.lbl_count.setText(
            f"{len(items)} articoli  ({n_com} commerciali, {n_nor} normalizzati)"
        )

        # Aggiorna filtri combo (per riflettere nuovi dati)
        self._populate_filters()

    # ==================================================================
    #  Selezione e interazione albero
    # ==================================================================

    def _on_tree_item_changed(self, current: QTreeWidgetItem | None, _prev):
        if current is None:
            self._detail.clear()
            return
        role_type = current.data(0, _ROLE_TYPE)
        if role_type == "item":
            item_id = current.data(0, _ROLE_ID)
            self._detail.load(item_id)
            self.item_selected.emit(item_id)
        elif role_type in ("cat", "sub"):
            type_label = "Categoria" if role_type == "cat" else "Sottogruppo"
            self._detail.load_group(type_label, current.text(0))
        else:
            self._detail.clear()

    def _on_double_click(self, item: QTreeWidgetItem, col: int):
        if item.data(0, _ROLE_TYPE) == "item":
            self._action_open(item.data(0, _ROLE_ID))

    def _selected_item_id(self) -> int | None:
        cur = self.tree.currentItem()
        if cur and cur.data(0, _ROLE_TYPE) == "item":
            return cur.data(0, _ROLE_ID)
        return None

    # ==================================================================
    #  Context menu
    # ==================================================================

    def _context_menu(self, pos):
        item_id = self._selected_item_id()
        menu = QMenu(self)

        act_new = QAction("+ Nuovo articolo", self)
        act_new.triggered.connect(self._action_new_item)
        menu.addAction(act_new)

        if item_id:
            menu.addSeparator()
            act_open = QAction("Apri / Modifica", self)
            act_open.triggered.connect(lambda: self._action_open(item_id))
            menu.addAction(act_open)

            # Checkout / Checkin
            act_co = QAction("Checkout", self)
            act_co.triggered.connect(lambda: self._action_checkout(item_id))
            act_ci = QAction("Check-in", self)
            act_ci.triggered.connect(lambda: self._action_checkin(item_id))
            act_undo = QAction("Annulla checkout", self)
            act_undo.triggered.connect(lambda: self._action_undo_checkout(item_id))
            menu.addAction(act_co)
            menu.addAction(act_ci)
            menu.addAction(act_undo)

            menu.addSeparator()
            act_dup = QAction("Duplica articolo", self)
            act_dup.triggered.connect(lambda: self._action_duplicate(item_id))
            menu.addAction(act_dup)

            menu.addSeparator()
            act_state = QAction("Cambia stato…", self)
            act_state.triggered.connect(lambda: self._action_change_state(item_id))
            menu.addAction(act_state)

            act_del = QAction("Elimina", self)
            act_del.triggered.connect(lambda: self._action_delete(item_id))
            menu.addAction(act_del)

        menu.exec(self.tree.viewport().mapToGlobal(pos))

    # ==================================================================
    #  Azioni
    # ==================================================================

    def _action_new_item(self):
        from ui.commercial_item_dialog import CommercialItemDialog
        dlg = CommercialItemDialog(parent=self)
        dlg.saved.connect(lambda _: self.refresh())
        dlg.exec()

    def _action_open(self, item_id: int):
        from ui.commercial_item_dialog import CommercialItemDialog
        dlg = CommercialItemDialog(item_id=item_id, parent=self)
        dlg.saved.connect(lambda _: self.refresh())
        dlg.exec()

    def _action_duplicate(self, item_id: int):
        src = session.commercial.get_item(item_id)
        if not src:
            return
        from ui.commercial_item_dialog import CommercialItemDialog
        dlg = CommercialItemDialog(prefill=dict(src), parent=self)
        dlg.saved.connect(lambda _: self.refresh())
        dlg.exec()

    def _action_checkout(self, item_id: int):
        if not session.sp:
            return
        try:
            ws_path = session.commercial.checkout_item(
                item_id, session.user, session.sp
            )
            QMessageBox.information(
                self, "Checkout completato",
                f"File nella workspace:\n{ws_path}"
            )
            self.refresh()
        except Exception as e:
            QMessageBox.critical(self, "Errore checkout", str(e))

    def _action_checkin(self, item_id: int):
        if not session.sp:
            return
        try:
            result = session.commercial.checkin_item(
                item_id, session.user, session.sp
            )
            msg = "Check-in completato."
            if result.get("conflict"):
                msg += "\n⚠ Conflitto: il file in archivio è stato modificato da un altro utente."
            QMessageBox.information(self, "Check-in", msg)
            self.refresh()
        except Exception as e:
            QMessageBox.critical(self, "Errore check-in", str(e))

    def _action_undo_checkout(self, item_id: int):
        reply = QMessageBox.question(
            self, "Annulla checkout",
            "Annullare il checkout? Le modifiche locali verranno perse.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        try:
            session.commercial.undo_checkout_item(item_id, session.user)
            self.refresh()
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

    def _action_change_state(self, item_id: int):
        item = session.commercial.get_item(item_id)
        if not item:
            return
        current_state = item["state"]
        available = COMMERCIAL_WORKFLOW_TRANSITIONS.get(current_state, [])
        if not available:
            QMessageBox.information(
                self, "Nessuna transizione",
                f"Nessuna transizione disponibile dallo stato '{current_state}'."
            )
            return

        from PyQt6.QtWidgets import QInputDialog
        target, ok = QInputDialog.getItem(
            self, "Cambia stato",
            f"Stato corrente: {current_state}\nNuovo stato:",
            available, 0, False,
        )
        if not ok:
            return
        try:
            session.commercial.change_state(
                item_id, target, session.user["id"]
            )
            self.refresh()
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

    def _action_delete(self, item_id: int):
        item = session.commercial.get_item(item_id)
        if not item:
            return
        reply = QMessageBox.question(
            self, "Conferma eliminazione",
            f"Eliminare l'articolo '{item['code']}'?\n"
            "L'operazione non può essere annullata.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        try:
            session.commercial.delete_item(item_id)
            self._detail.clear()
            self.refresh()
        except Exception as e:
            QMessageBox.critical(self, "Errore eliminazione", str(e))

    # ==================================================================
    #  Persistenza layout
    # ==================================================================

    def save_layout(self):
        settings = QSettings(_APP, _KEY)
        settings.setValue("splitter_state", self._splitter.saveState())


# =============================================================================
#  CommercialDetailPanel  –  Pannello informazioni articolo selezionato
# =============================================================================

class CommercialDetailPanel(QWidget):
    """Pannello di dettaglio (sola lettura) dell'articolo commerciale selezionato."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._build_ui()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(8)

        # ── Vista gruppo (categoria / sottogruppo) ──────────────────────
        self._grp_group = QFrame()
        grp_lay = QVBoxLayout(self._grp_group)
        grp_lay.setContentsMargins(4, 4, 4, 4)
        grp_lay.setSpacing(4)

        self.lbl_group_type = QLabel()
        font_gt = self.lbl_group_type.font()
        font_gt.setItalic(True)
        self.lbl_group_type.setFont(font_gt)
        self.lbl_group_type.setStyleSheet("color: #888;")

        self.lbl_group_code_desc = QLabel()
        self.lbl_group_code_desc.setWordWrap(True)
        self.lbl_group_code_desc.setFont(QFont("Segoe UI", 13, QFont.Weight.Bold))

        grp_lay.addWidget(self.lbl_group_type)
        grp_lay.addWidget(self.lbl_group_code_desc)
        grp_lay.addStretch()
        self._grp_group.hide()
        root.addWidget(self._grp_group)

        # ── Vista articolo ──────────────────────────────────────────────
        self._grp_item = QWidget()
        item_lay = QVBoxLayout(self._grp_item)
        item_lay.setContentsMargins(0, 0, 0, 0)
        item_lay.setSpacing(8)

        # Codice
        self.lbl_code = QLabel("—")
        self.lbl_code.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        self.lbl_code.setStyleSheet("color: #89b4fa;")
        self.lbl_code.setWordWrap(True)

        # Descrizione
        self.lbl_desc = QLabel()
        self.lbl_desc.setWordWrap(True)

        item_lay.addWidget(self.lbl_code)
        item_lay.addWidget(self.lbl_desc)

        # Informazioni
        self.lbl_state = QLabel()
        self.lbl_type  = QLabel()
        self.lbl_cat   = QLabel()
        self.lbl_cat.setWordWrap(True)
        info_grp = QGroupBox("Informazioni")
        info_form = QFormLayout(info_grp)
        info_form.addRow("Stato:", self.lbl_state)
        info_form.addRow("Tipo:", self.lbl_type)
        info_form.addRow("Categoria:", self.lbl_cat)
        item_lay.addWidget(info_grp)

        # Fornitori
        sup_grp = QGroupBox("Fornitori")
        sup_lay = QVBoxLayout(sup_grp)
        self.tbl_sup = QTableWidget(0, 4)
        self.tbl_sup.setHorizontalHeaderLabels(
            ["Fornitore", "Codice", "Prezzo", "Lead (gg)"]
        )
        self.tbl_sup.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeMode.Stretch
        )
        for col in (1, 2, 3):
            self.tbl_sup.horizontalHeader().setSectionResizeMode(
                col, QHeaderView.ResizeMode.ResizeToContents
            )
        self.tbl_sup.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tbl_sup.verticalHeader().setVisible(False)
        sup_lay.addWidget(self.tbl_sup)
        item_lay.addWidget(sup_grp)

        # Proprietà
        prop_grp = QGroupBox("Proprietà")
        prop_lay = QVBoxLayout(prop_grp)
        self.tbl_props = QTableWidget(0, 2)
        self.tbl_props.setHorizontalHeaderLabels(["Proprietà", "Valore"])
        self.tbl_props.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeMode.ResizeToContents
        )
        self.tbl_props.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeMode.Stretch
        )
        self.tbl_props.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tbl_props.verticalHeader().setVisible(False)
        prop_lay.addWidget(self.tbl_props)
        item_lay.addWidget(prop_grp)

        # Bottone Modifica
        btn_row = QHBoxLayout()
        btn_open = QPushButton("Modifica")
        btn_open.clicked.connect(self._open_item)
        btn_row.addWidget(btn_open)
        btn_row.addStretch()
        item_lay.addLayout(btn_row)

        self._grp_item.hide()
        root.addWidget(self._grp_item)
        root.addStretch()

        self._current_id: int | None = None

    # ------------------------------------------------------------------

    def load_group(self, type_label: str, code_desc: str):
        """Mostra informazioni di una categoria o sottogruppo."""
        self._current_id = None
        self._grp_item.hide()
        self.lbl_group_type.setText(type_label)
        self.lbl_group_code_desc.setText(code_desc)
        self._grp_group.show()

    def load(self, item_id: int):
        self._current_id = item_id
        item = session.commercial.get_item(item_id)
        if not item:
            self.clear()
            return

        self._grp_group.hide()

        self.lbl_code.setText(item.get("code") or "—")
        self.lbl_desc.setText(item.get("description") or "")
        self.lbl_state.setText(item.get("state") or "")
        type_label = "Commerciale (5)" if item["item_type"] == "commerciale" else "Normalizzato (6)"
        self.lbl_type.setText(type_label)
        cat_text = item.get("cat_description") or ""
        if item.get("sub_description"):
            cat_text += f" / {item['sub_description']}"
        self.lbl_cat.setText(cat_text)

        # Fornitori
        self.tbl_sup.setRowCount(0)
        for lnk in session.commercial.get_item_suppliers(item_id):
            r = self.tbl_sup.rowCount()
            self.tbl_sup.insertRow(r)
            price = f"{lnk['unit_price']:.4f} {lnk.get('currency','EUR')}" \
                if lnk.get("unit_price") is not None else "—"
            lead = str(lnk["lead_time_days"]) if lnk.get("lead_time_days") else "—"
            name = lnk.get("supplier_name") or ""
            if lnk.get("is_preferred"):
                name = "★ " + name
            self.tbl_sup.setItem(r, 0, QTableWidgetItem(name))
            self.tbl_sup.setItem(r, 1, QTableWidgetItem(lnk.get("supplier_code") or ""))
            self.tbl_sup.setItem(r, 2, QTableWidgetItem(price))
            self.tbl_sup.setItem(r, 3, QTableWidgetItem(lead))

        # Proprietà
        self.tbl_props.setRowCount(0)
        props = session.commercial.get_properties(item_id)
        for name, value in props.items():
            r = self.tbl_props.rowCount()
            self.tbl_props.insertRow(r)
            self.tbl_props.setItem(r, 0, QTableWidgetItem(name))
            self.tbl_props.setItem(r, 1, QTableWidgetItem(str(value) if value is not None else ""))

        self._grp_item.show()

    def clear(self):
        self._current_id = None
        self._grp_group.hide()
        self._grp_item.hide()
        self.tbl_sup.setRowCount(0)
        self.tbl_props.setRowCount(0)

    def _open_item(self):
        if self._current_id is None:
            return
        from ui.commercial_item_dialog import CommercialItemDialog
        dlg = CommercialItemDialog(item_id=self._current_id, parent=self)
        dlg.exec()

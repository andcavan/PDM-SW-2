# =============================================================================
#  ui/archive_view.py  –  Vista archivio CAD (QTreeWidget raggruppato per codice)
# =============================================================================
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QComboBox, QTreeWidget, QTreeWidgetItem,
    QHeaderView, QMenu, QMessageBox, QFileDialog, QSplitter
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QColor, QAction
from pathlib import Path

from config import WORKFLOW_STATES
from core.checkout_manager import READONLY_STATES
from core.file_manager import EXT_FOR_TYPE
from ui.session import session
from ui.styles import TYPE_ICON


# Colonne del tree
COL_CODE   = 0
COL_REV    = 1
COL_TITLE  = 2
COL_STATE  = 3
COL_LOCK   = 4
COL_AUTHOR = 5
COL_DATE   = 6


class ArchiveView(QWidget):
    """
    Vista archivio CAD con struttura ad albero:
    - Nodi padre  = Codice (raggruppa tutte le revisioni/tipi)
    - Nodi figlio = singolo documento (Parte/Assieme/Disegno con revisione)
    Il filtro "Tipo" agisce solo sui figli: nasconde i tipi non corrispondenti.
    """

    document_selected = pyqtSignal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._build_ui()

    # ------------------------------------------------------------------
    #  UI
    # ------------------------------------------------------------------
    def _build_ui(self):
        from ui.detail_panel import DetailPanel

        layout = QVBoxLayout(self)
        layout.setSpacing(6)
        layout.setContentsMargins(8, 8, 8, 8)

        # Barra filtri
        flt = QHBoxLayout()

        flt.addWidget(QLabel("Cerca:"))
        self.txt_search = QLineEdit()
        self.txt_search.setPlaceholderText("Codice, titolo, descrizione...")
        self.txt_search.setClearButtonEnabled(True)
        self.txt_search.returnPressed.connect(self.refresh)
        flt.addWidget(self.txt_search, 2)

        flt.addWidget(QLabel("Tipo:"))
        self.cmb_type = QComboBox()
        self.cmb_type.addItems(["Tutti", "Parte", "Assieme", "Disegno"])
        self.cmb_type.currentIndexChanged.connect(self._apply_type_filter)
        flt.addWidget(self.cmb_type)

        flt.addWidget(QLabel("Stato:"))
        self.cmb_state = QComboBox()
        self.cmb_state.addItems(["Tutti"] + list(WORKFLOW_STATES.keys()))
        self.cmb_state.currentIndexChanged.connect(self.refresh)
        flt.addWidget(self.cmb_state)

        btn_refresh = QPushButton("\u21bb  Aggiorna")
        btn_refresh.clicked.connect(self.refresh)
        flt.addWidget(btn_refresh)

        layout.addLayout(flt)

        # ---- Splitter: albero | pannello dettaglio ----
        self.splitter = QSplitter(Qt.Orientation.Horizontal)

        # Lato sinistro: tree + conteggio
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(4)

        # Barra espandi/comprimi
        exp_row = QHBoxLayout()
        btn_expand = QPushButton("▶  Espandi tutto")
        btn_expand.setToolTip("Espandi tutti i nodi codice")
        btn_expand.clicked.connect(self.tree_expand_all)
        btn_collapse = QPushButton("◀  Comprimi tutto")
        btn_collapse.setToolTip("Comprimi tutti i nodi codice")
        btn_collapse.clicked.connect(self.tree_collapse_all)
        exp_row.addWidget(btn_collapse)
        exp_row.addWidget(btn_expand)
        exp_row.addStretch()
        left_layout.addLayout(exp_row)

        # Tree widget
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels([
            "Codice", "Rev.", "Titolo", "Stato",
            "Checkout", "Creato da", "Data mod."
        ])
        hdr = self.tree.header()
        hdr.setSectionResizeMode(COL_CODE,   QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(COL_REV,    QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(COL_TITLE,  QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(COL_STATE,  QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(COL_LOCK,   QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(COL_AUTHOR, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(COL_DATE,   QHeaderView.ResizeMode.ResizeToContents)

        self.tree.setSelectionBehavior(QTreeWidget.SelectionBehavior.SelectRows)
        self.tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self._context_menu)
        self.tree.setRootIsDecorated(True)
        self.tree.setAnimated(True)
        self.tree.setAlternatingRowColors(True)
        self.tree.itemDoubleClicked.connect(self._on_double_click)
        self.tree.itemSelectionChanged.connect(self._on_selection_changed)

        left_layout.addWidget(self.tree)

        # Conteggio
        self.lbl_count = QLabel("")
        self.lbl_count.setObjectName("subtitle_label")
        left_layout.addWidget(self.lbl_count)

        self.splitter.addWidget(left_widget)

        # Lato destro: pannello dettaglio
        self.detail_panel = DetailPanel()
        self.splitter.addWidget(self.detail_panel)

        # Connetti segnali creazione archive-first
        self.detail_panel.create_in_sw_requested.connect(self._action_create_in_sw)
        self.detail_panel.create_from_file_requested.connect(self._action_create_from_file)
        self.detail_panel.add_drw_requested.connect(self._action_add_drw)

        # Proporzioni iniziali 65/35
        self.splitter.setStretchFactor(0, 65)
        self.splitter.setStretchFactor(1, 35)
        self.splitter.setCollapsible(0, False)
        self.splitter.setCollapsible(1, True)

        # Ripristina dimensioni splitter salvate
        s = QSettings("PDM-SW", "ArchiveView")
        splitter_state = s.value("splitterState")
        if splitter_state:
            self.splitter.restoreState(splitter_state)

        layout.addWidget(self.splitter)

    def save_layout(self):
        """Salva lo stato dello splitter (chiamata da MainWindow.closeEvent)."""
        s = QSettings("PDM-SW", "ArchiveView")
        s.setValue("splitterState", self.splitter.saveState())

    # ------------------------------------------------------------------
    #  Refresh dati
    # ------------------------------------------------------------------
    def refresh(self):
        if not session.is_connected:
            return

        text  = self.txt_search.text().strip()
        state = self.cmb_state.currentText()
        if state == "Tutti":
            state = ""

        docs = session.files.search_documents(text=text, state=state)

        # Ricorda quali codici erano espansi e la selezione corrente
        expanded_codes: set[str] = set()
        selected_key = None
        root_item = self.tree.invisibleRootItem()
        for i in range(root_item.childCount()):
            child = root_item.child(i)
            if child.isExpanded():
                val = child.data(COL_CODE, Qt.ItemDataRole.UserRole)
                if isinstance(val, str) and val.startswith("CODE:"):
                    expanded_codes.add(val[5:])
        sel_items = self.tree.selectedItems()
        if sel_items:
            selected_key = sel_items[0].data(COL_CODE, Qt.ItemDataRole.UserRole)

        # Raggruppa per codice
        groups: dict[str, list[dict]] = {}
        for d in docs:
            groups.setdefault(d["code"], []).append(d)

        self.tree.clear()

        for code in sorted(groups.keys()):
            all_docs = groups[code]
            parent_item = QTreeWidgetItem(self.tree)
            parent_item.setText(COL_CODE, code)
            # Nodo codice: selezionabile, memorizza il codice stringa
            parent_item.setData(COL_CODE, Qt.ItemDataRole.UserRole, f"CODE:{code}")

            # Titolo dal documento rappresentativo
            best = self._pick_representative(all_docs)
            parent_item.setText(COL_TITLE, best.get("title") or "")

            # Figli visibili solo se hanno un file in archivio O sono in checkout
            visible_docs = [d for d in all_docs
                            if d.get("archive_path") or d["is_locked"]]

            # Separa doc attivi e obsoleti tra i visibili
            active_docs   = [d for d in visible_docs if d["state"] != "Obsoleto"]
            obsolete_docs = [d for d in visible_docs if d["state"] == "Obsoleto"]

            # Crea nodi per documenti attivi, tracciando l'ultimo per tipo
            type_items: dict[str, QTreeWidgetItem] = {}
            for doc in sorted(active_docs, key=lambda d: (d["doc_type"], d["revision"])):
                item = self._make_tree_item(parent_item, doc)
                type_items[doc["doc_type"]] = item

            # Annida obsoleti sotto l'ultimo doc attivo dello stesso tipo
            for doc in sorted(obsolete_docs,
                              key=lambda d: (d["doc_type"], d["revision"]),
                              reverse=True):
                nest_parent = type_items.get(doc["doc_type"], parent_item)
                self._make_tree_item(nest_parent, doc)

            parent_item.setExpanded(code in expanded_codes)   # ripristina stato o lascia compresso

        total = sum(len(v) for v in groups.values())
        self.lbl_count.setText(
            f"{total} documenti in {len(groups)} codici"
        )

        # Applica filtro tipo corrente
        self._apply_type_filter()

        # Ripristina selezione precedente (senza emettere segnali extra)
        if selected_key is not None:
            root = self.tree.invisibleRootItem()
            found: "QTreeWidgetItem | None" = None
            for i in range(root.childCount()):
                parent = root.child(i)
                if parent.data(COL_CODE, Qt.ItemDataRole.UserRole) == selected_key:
                    found = parent
                    break
                for j in range(parent.childCount()):
                    child = parent.child(j)
                    if child.data(COL_CODE, Qt.ItemDataRole.UserRole) == selected_key:
                        found = child
                        break
                if found:
                    break
            if found:
                self.tree.blockSignals(True)
                self.tree.setCurrentItem(found)
                self.tree.blockSignals(False)
                # Aggiorna il pannello dettaglio con i dati freschi
                self._on_selection_changed()

    # ------------------------------------------------------------------
    #  Espandi / Comprimi
    # ------------------------------------------------------------------
    def tree_expand_all(self):
        self.tree.expandAll()

    def tree_collapse_all(self):
        root = self.tree.invisibleRootItem()
        for i in range(root.childCount()):
            root.child(i).setExpanded(False)

    # ------------------------------------------------------------------
    #  Filtro tipo (agisce solo sui figli)
    # ------------------------------------------------------------------
    def _apply_type_filter(self):
        chosen = self.cmb_type.currentText()
        root = self.tree.invisibleRootItem()
        for i in range(root.childCount()):
            parent = root.child(i)
            parent.setHidden(False)   # i nodi codice sono sempre visibili
            for j in range(parent.childCount()):
                child = parent.child(j)
                if chosen == "Tutti":
                    child.setHidden(False)
                else:
                    text = child.text(COL_CODE).strip()
                    child.setHidden(chosen not in text)

    # ------------------------------------------------------------------
    #  Double-click
    # ------------------------------------------------------------------
    def _on_selection_changed(self):
        """Aggiorna il pannello dettaglio al cambio selezione."""
        items = self.tree.selectedItems()
        if not items:
            self.detail_panel.clear()
            return
        item = items[0]
        val = item.data(COL_CODE, Qt.ItemDataRole.UserRole)
        if isinstance(val, str) and val.startswith("CODE:"):
            code = val[5:]
            all_docs_raw = session.files.search_documents(code=code)
            all_docs = [d for d in all_docs_raw if d["code"] == code]
            self.detail_panel.load_code(code, all_docs)
        elif isinstance(val, int):
            self.detail_panel.load_document(val)
        else:
            self.detail_panel.clear()

    def _on_double_click(self, item: QTreeWidgetItem, column: int):
        val = item.data(COL_CODE, Qt.ItemDataRole.UserRole)
        if isinstance(val, int):
            self.document_selected.emit(val)

    # ------------------------------------------------------------------
    #  Helpers
    # ------------------------------------------------------------------
    @staticmethod
    def _pick_representative(docs: list[dict]) -> dict:
        """Sceglie il doc 'rappresentativo' di un codice (revisione piu alta, non obsoleto)."""
        active = [d for d in docs if d["state"] != "Obsoleto"]
        pool = active if active else docs
        return max(pool, key=lambda d: d["revision"])

    def _make_tree_item(self, parent: QTreeWidgetItem, doc: dict) -> QTreeWidgetItem:
        """Crea un nodo albero per un documento."""
        child = QTreeWidgetItem(parent)
        icon = TYPE_ICON.get(doc["doc_type"], "")
        child.setText(COL_CODE, f"  {icon}  {doc['doc_type']}")
        child.setText(COL_REV, doc["revision"])
        child.setText(COL_TITLE, doc.get("title") or "")

        st = doc["state"]
        child.setText(COL_STATE, st)
        color = WORKFLOW_STATES.get(st, {}).get("color", "#9E9E9E")
        child.setForeground(COL_STATE, QColor(color))

        if doc["is_locked"]:
            locker = doc.get("locked_by_name") or "?"
            child.setText(COL_LOCK, f"\U0001f512 {locker}")
            child.setForeground(COL_LOCK, QColor("#fab387"))
        else:
            child.setText(COL_LOCK, "")

        child.setText(COL_AUTHOR, doc.get("created_by_name") or "")
        mod_at = doc.get("modified_at") or doc.get("created_at") or ""
        child.setText(COL_DATE, str(mod_at)[:19] if mod_at else "")
        child.setData(COL_CODE, Qt.ItemDataRole.UserRole, doc["id"])
        return child

    def _selected_doc_id(self) -> int | None:
        """Ritorna il document_id del nodo figlio selezionato, o None."""
        items = self.tree.selectedItems()
        if not items:
            return None
        item = items[0]
        val = item.data(COL_CODE, Qt.ItemDataRole.UserRole)
        return val if isinstance(val, int) else None

    def _selected_doc(self) -> dict | None:
        doc_id = self._selected_doc_id()
        if not doc_id:
            return None
        return session.files.get_document(doc_id)

    # ------------------------------------------------------------------
    #  Menu contestuale
    # ------------------------------------------------------------------
    def _context_menu(self, pos):
        doc_id = self._selected_doc_id()
        if not doc_id:
            return
        doc = session.files.get_document(doc_id)
        if not doc:
            return

        menu = QMenu(self)
        is_readonly = doc["state"] in READONLY_STATES
        is_locked   = doc["is_locked"]
        uid         = session.user["id"] if session.user else -1
        is_my_lock  = is_locked and doc.get("locked_by") == uid
        is_admin    = session.can("admin")
        is_latest   = session.workflow.is_latest_revision(doc)

        # ---- Consultazione (copia senza lock) ----
        act_consult = QAction("👁  Consultazione", self)
        act_consult.setEnabled(bool(doc.get("archive_path")))
        act_consult.triggered.connect(lambda: self._action_consultation(doc_id))
        menu.addAction(act_consult)

        menu.addSeparator()

        # ---- Checkout ----
        act_checkout = QAction("📤  Checkout", self)
        act_checkout.setEnabled(
            not is_readonly and not is_locked and session.can("checkout")
        )
        act_checkout.triggered.connect(lambda: self._action_checkout(doc_id))
        menu.addAction(act_checkout)

        # ---- Checkin ----
        act_checkin = QAction("📥  Check-in", self)
        act_checkin.setEnabled(is_my_lock and not is_readonly)
        act_checkin.triggered.connect(lambda: self._action_checkin(doc_id))
        menu.addAction(act_checkin)

        # ---- Annulla checkout ----
        act_undo = QAction("↩️  Annulla checkout", self)
        act_undo.setEnabled(is_locked and (is_my_lock or is_admin))
        act_undo.triggered.connect(lambda: self._action_undo_checkout(doc_id))
        menu.addAction(act_undo)

        menu.addSeparator()

        # ---- Crea revisione (solo Rilasciato, ultima rev) ----
        act_newrev = QAction("📋  Crea revisione", self)
        act_newrev.setEnabled(
            doc["state"] == "Rilasciato"
            and is_latest
            and session.can("create")
        )
        act_newrev.triggered.connect(lambda: self._action_new_revision(doc_id))
        menu.addAction(act_newrev)

        # ---- Annulla revisione (solo In Revisione) ----
        act_cancel_rev = QAction("🗑  Annulla revisione", self)
        act_cancel_rev.setEnabled(
            doc["state"] == "In Revisione"
            and not is_locked
            and (is_my_lock or is_admin or not is_locked)
        )
        act_cancel_rev.triggered.connect(lambda: self._action_cancel_revision(doc_id))
        menu.addAction(act_cancel_rev)

        menu.addSeparator()

        # ---- Workflow ----
        act_wf = QAction("🔄  Workflow", self)
        act_wf.setEnabled(is_latest and (not is_readonly or is_admin))
        act_wf.triggered.connect(lambda: self._action_workflow(doc_id))
        menu.addAction(act_wf)

        menu.addSeparator()

        # ---- Apri in eDrawings ----
        act_edraw = QAction("👁️  Apri in eDrawings", self)
        act_edraw.setEnabled(bool(doc.get("archive_path")))
        act_edraw.triggered.connect(lambda: self._action_edrawings(doc_id))
        menu.addAction(act_edraw)

        # ---- Proprietà ----
        act_props = QAction("ℹ️  Proprietà", self)
        act_props.triggered.connect(lambda: self._action_properties(doc_id))
        menu.addAction(act_props)

        menu.exec(self.tree.viewport().mapToGlobal(pos))

    # ------------------------------------------------------------------
    #  Azioni
    # ------------------------------------------------------------------
    def _action_checkout(self, doc_id: int):
        from ui.checkout_dialog import CheckoutDialog
        dlg = CheckoutDialog(doc_id, parent=self)
        if dlg.exec() == CheckoutDialog.DialogCode.Accepted:
            self.refresh()

    def _action_checkin(self, doc_id: int):
        from ui.checkin_dialog import CheckinDialog
        dlg = CheckinDialog(doc_id, parent=self)
        if dlg.exec() == CheckinDialog.DialogCode.Accepted:
            self.refresh()

    def _action_undo_checkout(self, doc_id: int):
        r = QMessageBox.question(
            self, "Annulla checkout",
            "Annullare il checkout senza salvare?\n"
            "Il file nella workspace non verra archiviato.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if r == QMessageBox.StandardButton.Yes:
            try:
                session.checkout.undo_checkout(doc_id, delete_from_workspace=True)
                QMessageBox.information(self, "OK", "Checkout annullato")
                self.refresh()
            except Exception as e:
                QMessageBox.critical(self, "Errore", str(e))

    def _action_consultation(self, doc_id: int):
        try:
            dest = session.checkout.open_for_consultation(doc_id)
            QMessageBox.information(
                self, "Consultazione",
                f"File copiato in workspace:\n{dest.name}\n\n"
                "Il file e in sola lettura (nessun lock)."
            )
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

    def _action_workflow(self, doc_id: int):
        from ui.workflow_dialog import WorkflowDialog
        dlg = WorkflowDialog(doc_id, parent=self)
        if dlg.exec():
            self.refresh()

    def _action_new_revision(self, doc_id: int):
        """Crea nuova revisione del documento."""
        doc = session.files.get_document(doc_id)
        if not doc:
            return

        # R3: la nuova revisione di un DRW va fatta dal PRT/ASM associato
        if doc["doc_type"] == "Disegno":
            QMessageBox.warning(
                self, "Operazione non consentita",
                "La nuova revisione di un Disegno deve essere creata\n"
                "dal PRT/ASM associato, non direttamente."
            )
            return

        # Determina prossima revisione (incremento numerico semplice)
        try:
            current_num = int(doc["revision"])
            next_rev = str(current_num + 1).zfill(len(doc["revision"]))
        except ValueError:
            # Revisione non numerica: chiedi all'utente
            from PyQt6.QtWidgets import QInputDialog
            next_rev, ok = QInputDialog.getText(
                self, "Nuova revisione",
                f"Revisione attuale: {doc['revision']}\nInserisci nuova revisione:"
            )
            if not ok or not next_rev.strip():
                return
            next_rev = next_rev.strip()

        # Verifica che la revisione non esista gia
        existing = session.db.fetchone(
            "SELECT id FROM documents WHERE code=? AND revision=?",
            (doc["code"], next_rev),
        )
        if existing:
            QMessageBox.warning(
                self, "Revisione esistente",
                f"La revisione {next_rev} esiste gia per il codice {doc['code']}."
            )
            return

        r = QMessageBox.question(
            self, "Nuova revisione",
            f"Creare revisione {next_rev} dal codice {doc['code']} rev. {doc['revision']}?\n\n"
            "Il file archiviato verra copiato come base per la nuova revisione.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if r != QMessageBox.StandardButton.Yes:
            return

        try:
            new_id = session.workflow.new_revision(
                doc_id, session.user["id"], next_rev,
                shared_paths=session.sp,
            )
            QMessageBox.information(
                self, "Nuova revisione",
                f"Creata revisione {next_rev} (ID: {new_id})"
            )
            self.refresh()
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

    def _action_cancel_revision(self, doc_id: int):
        """Annulla revisione In Revisione — elimina il documento e torna alla precedente."""
        doc = session.files.get_document(doc_id)
        if not doc:
            return
        r = QMessageBox.question(
            self, "Annulla revisione",
            f"Annullare la revisione {doc['revision']} del codice {doc['code']}?\n\n"
            "Il documento e il relativo file archiviato verranno eliminati.\n"
            "La revisione precedente resterà inalterata.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if r != QMessageBox.StandardButton.Yes:
            return
        try:
            session.workflow.cancel_revision(
                doc_id, session.user["id"],
                shared_paths=session.sp,
            )
            QMessageBox.information(
                self, "Annulla revisione",
                f"Revisione {doc['revision']} annullata."
            )
            self.refresh()
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

    def _action_edrawings(self, doc_id: int):
        """Apre il file in eDrawings per consultazione rapida."""
        self.detail_panel.load_document(doc_id)
        self.detail_panel._open_edrawings()

    def _action_properties(self, doc_id: int):
        from ui.document_dialog import DocumentDialog
        dlg = DocumentDialog(doc_id, parent=self)
        if dlg.exec():
            self.refresh()

    # ------------------------------------------------------------------
    #  Azioni creazione archive-first (attivate da segnali DetailPanel)
    # ------------------------------------------------------------------
    def _action_create_in_sw(self, doc_id: int):
        """Crea il template SW direttamente in archivio; optional checkout."""
        also_checkout = self.detail_panel.chk_checkout.isChecked()
        try:
            dest = session.files.create_to_archive(doc_id)
            if also_checkout:
                session.checkout.checkout(doc_id)
                session.files.open_from_workspace(doc_id)
            msg = f"File creato in archivio:\n{dest.name}"
            if also_checkout:
                msg += "\n\nFile aperto in SolidWorks."
            QMessageBox.information(self, "File creato", msg)
            self.refresh()
        except FileNotFoundError as e:
            r = QMessageBox.question(
                self, "Template non configurato",
                f"{e}\n\nAprire la configurazione SolidWorks ora?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if r == QMessageBox.StandardButton.Yes:
                from ui.sw_config_dialog import SWConfigDialog
                dlg = SWConfigDialog(parent=self)
                dlg.exec()
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

    def _action_create_from_file(self, doc_id: int):
        """Importa un file esterno direttamente in archivio."""
        doc = session.files.get_document(doc_id)
        if not doc:
            return
        sw_ext = EXT_FOR_TYPE.get(doc["doc_type"], ".SLDPRT")
        path, _ = QFileDialog.getOpenFileName(
            self, f"Seleziona file {doc['doc_type']}",
            "", f"File SolidWorks (*{sw_ext} *{sw_ext.lower()})"
        )
        if not path:
            return
        try:
            dest = session.files.create_to_archive(doc_id, source_path=Path(path))
            also_checkout = self.detail_panel.chk_checkout.isChecked()
            if also_checkout:
                session.checkout.checkout(doc_id)
            QMessageBox.information(
                self, "File importato",
                f"File importato in archivio:\n{dest.name}"
            )
            self.refresh()
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

    def _action_add_drw(self, parent_doc_id: int):
        """Crea il disegno DRW in archivio partendo dal template."""
        try:
            drw_id = session.files.get_or_create_drw_document(parent_doc_id)
            dest = session.files.create_to_archive(drw_id)
            also_checkout = self.detail_panel.chk_checkout.isChecked()
            if also_checkout:
                session.checkout.checkout(drw_id)
            QMessageBox.information(
                self, "DRW creato",
                f"Disegno creato in archivio:\n{dest.name}"
            )
            self.refresh()
        except FileNotFoundError as e:
            r = QMessageBox.question(
                self, "Template non configurato",
                f"{e}\n\nAprire la configurazione SolidWorks ora?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if r == QMessageBox.StandardButton.Yes:
                from ui.sw_config_dialog import SWConfigDialog
                dlg = SWConfigDialog(parent=self)
                dlg.exec()
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

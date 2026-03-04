# =============================================================================
#  ui/archive_view.py  –  Vista archivio CAD (QTreeWidget raggruppato per codice)
# =============================================================================
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QComboBox, QTreeWidget, QTreeWidgetItem,
    QHeaderView, QMenu, QMessageBox, QFileDialog
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QColor, QAction

from config import WORKFLOW_STATES
from core.checkout_manager import READONLY_STATES
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

        layout.addWidget(self.tree)

        # Conteggio
        self.lbl_count = QLabel("")
        self.lbl_count.setObjectName("subtitle_label")
        layout.addWidget(self.lbl_count)

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

        # Raggruppa per codice
        groups: dict[str, list[dict]] = {}
        for d in docs:
            groups.setdefault(d["code"], []).append(d)

        self.tree.clear()

        for code in sorted(groups.keys()):
            children = groups[code]
            parent_item = QTreeWidgetItem(self.tree)
            parent_item.setText(COL_CODE, code)
            parent_item.setFlags(
                parent_item.flags() & ~Qt.ItemFlag.ItemIsSelectable
            )

            # Titolo e stato dal primo figlio (o quello con revisione piu alta)
            best = self._pick_representative(children)
            parent_item.setText(COL_TITLE, best.get("title") or "")
            parent_item.setData(COL_CODE, Qt.ItemDataRole.UserRole, None)

            # Figli: un nodo per ogni documento (tipo + revisione)
            for doc in sorted(children, key=lambda d: (d["doc_type"], d["revision"])):
                child = QTreeWidgetItem(parent_item)
                icon = TYPE_ICON.get(doc["doc_type"], "")
                child.setText(COL_CODE, f"  {icon}  {doc['doc_type']}")
                child.setText(COL_REV, doc["revision"])
                child.setText(COL_TITLE, doc.get("title") or "")

                # Stato con colore
                st = doc["state"]
                child.setText(COL_STATE, st)
                color = WORKFLOW_STATES.get(st, {}).get("color", "#9E9E9E")
                child.setForeground(COL_STATE, QColor(color))

                # Checkout
                if doc["is_locked"]:
                    locker = doc.get("locked_by_name") or "?"
                    child.setText(COL_LOCK, f"\U0001f512 {locker}")
                    child.setForeground(COL_LOCK, QColor("#fab387"))
                else:
                    child.setText(COL_LOCK, "")

                child.setText(COL_AUTHOR, doc.get("created_by_name") or "")

                mod_at = doc.get("modified_at") or doc.get("created_at") or ""
                child.setText(COL_DATE, str(mod_at)[:19] if mod_at else "")

                # Store doc id
                child.setData(COL_CODE, Qt.ItemDataRole.UserRole, doc["id"])

            parent_item.setExpanded(True)

        total = sum(len(v) for v in groups.values())
        self.lbl_count.setText(
            f"{total} documenti in {len(groups)} codici"
        )

        # Applica filtro tipo corrente
        self._apply_type_filter()

    # ------------------------------------------------------------------
    #  Filtro tipo (agisce solo sui figli)
    # ------------------------------------------------------------------
    def _apply_type_filter(self):
        chosen = self.cmb_type.currentText()
        root = self.tree.invisibleRootItem()
        for i in range(root.childCount()):
            parent = root.child(i)
            any_visible = False
            for j in range(parent.childCount()):
                child = parent.child(j)
                if chosen == "Tutti":
                    child.setHidden(False)
                    any_visible = True
                else:
                    # Il testo colonna 0 del figlio contiene l'icona + tipo
                    text = child.text(COL_CODE).strip()
                    visible = chosen in text
                    child.setHidden(not visible)
                    if visible:
                        any_visible = True
            parent.setHidden(not any_visible)

    # ------------------------------------------------------------------
    #  Double-click
    # ------------------------------------------------------------------
    def _on_double_click(self, item: QTreeWidgetItem, column: int):
        doc_id = item.data(COL_CODE, Qt.ItemDataRole.UserRole)
        if doc_id:
            self.document_selected.emit(doc_id)

    # ------------------------------------------------------------------
    #  Helpers
    # ------------------------------------------------------------------
    @staticmethod
    def _pick_representative(docs: list[dict]) -> dict:
        """Sceglie il doc 'rappresentativo' di un codice (revisione piu alta, non obsoleto)."""
        active = [d for d in docs if d["state"] != "Obsoleto"]
        pool = active if active else docs
        return max(pool, key=lambda d: d["revision"])

    def _selected_doc_id(self) -> int | None:
        """Ritorna il document_id del nodo figlio selezionato, o None."""
        items = self.tree.selectedItems()
        if not items:
            return None
        item = items[0]
        doc_id = item.data(COL_CODE, Qt.ItemDataRole.UserRole)
        return doc_id

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

        # ---- Checkout ----
        act_checkout = QAction("\U0001f4e4  Checkout", self)
        act_checkout.setEnabled(
            not is_readonly and not is_locked and session.can("checkout")
        )
        act_checkout.triggered.connect(lambda: self._action_checkout(doc_id))
        menu.addAction(act_checkout)

        # ---- Checkin ----
        act_checkin = QAction("\U0001f4e5  Check-in", self)
        act_checkin.setEnabled(is_my_lock and not is_readonly)
        act_checkin.triggered.connect(lambda: self._action_checkin(doc_id))
        menu.addAction(act_checkin)

        # ---- Annulla checkout ----
        act_undo = QAction("\u21a9\ufe0f  Annulla checkout", self)
        act_undo.setEnabled(is_locked and (is_my_lock or is_admin))
        act_undo.triggered.connect(lambda: self._action_undo_checkout(doc_id))
        menu.addAction(act_undo)

        menu.addSeparator()

        # ---- Consultazione (copia senza lock) ----
        act_consult = QAction("\U0001f441  Consultazione", self)
        act_consult.setEnabled(bool(doc.get("archive_path")))
        act_consult.triggered.connect(lambda: self._action_consultation(doc_id))
        menu.addAction(act_consult)

        # ---- Esporta da workspace ----
        act_export = QAction("\U0001f4c2  Esporta da workspace", self)
        act_export.triggered.connect(lambda: self._action_export(doc_id))
        menu.addAction(act_export)

        menu.addSeparator()

        # ---- Workflow ----
        act_wf = QAction("\U0001f504  Cambia stato", self)
        act_wf.setEnabled(not is_readonly or is_admin)
        act_wf.triggered.connect(lambda: self._action_workflow(doc_id))
        menu.addAction(act_wf)

        # ---- Nuova revisione ----
        act_newrev = QAction("\U0001f4cb  Nuova revisione", self)
        act_newrev.setEnabled(
            doc["state"] in ("Rilasciato", "Revisionato")
            and session.can("create")
        )
        act_newrev.triggered.connect(lambda: self._action_new_revision(doc_id))
        menu.addAction(act_newrev)

        menu.addSeparator()

        # ---- Proprieta ----
        act_props = QAction("\u2139\ufe0f  Proprieta", self)
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

    def _action_export(self, doc_id: int):
        folder = QFileDialog.getExistingDirectory(
            self, "Seleziona cartella di destinazione"
        )
        if not folder:
            return
        try:
            from pathlib import Path
            exported = session.files.export_from_workspace(doc_id, Path(folder))
            names = ", ".join(p.name for p in exported)
            QMessageBox.information(
                self, "Esportazione", f"Esportati:\n{names}"
            )
        except FileNotFoundError as e:
            QMessageBox.warning(self, "File non in workspace", str(e))
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

    def _action_properties(self, doc_id: int):
        from ui.document_dialog import DocumentDialog
        dlg = DocumentDialog(doc_id, parent=self)
        if dlg.exec():
            self.refresh()

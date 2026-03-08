# =============================================================================
#  ui/workspace_view.py  –  Vista workspace  v3.1
#  Scansiona la cartella workspace e mostra stato PDM per ogni file
# =============================================================================
from pathlib import Path
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
    QComboBox, QMessageBox
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor

from config import WORKFLOW_STATES
from ui.session import session
from ui.styles import TYPE_ICON

# Estensioni SolidWorks → tipo documento PDM
_EXT_TO_TYPE = {
    ".sldprt": "Parte",
    ".sldasm": "Assieme",
    ".slddrw": "Disegno",
}

_ROLE_DOC_ID      = Qt.ItemDataRole.UserRole
_ROLE_WS_PATH     = Qt.ItemDataRole.UserRole + 1
_ROLE_WS_STATUS   = Qt.ItemDataRole.UserRole + 2
_ROLE_SEARCH_TEXT = Qt.ItemDataRole.UserRole + 3   # codice + titolo + descrizione (minuscolo)
_ROLE_CO_STATE    = Qt.ItemDataRole.UserRole + 4   # "mine" | "other" | "free" | "unknown"


class WorkspaceView(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._build_ui()

    # ------------------------------------------------------------------
    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(8)
        layout.setContentsMargins(8, 8, 8, 8)

        # ---- Header ----
        hdr = QHBoxLayout()
        self.lbl_user = QLabel("—")
        self.lbl_user.setObjectName("title_label")
        hdr.addWidget(self.lbl_user)
        hdr.addStretch()
        self.lbl_ws_path = QLabel("")
        self.lbl_ws_path.setStyleSheet("color:#6c7086;font-size:11px;")
        hdr.addWidget(self.lbl_ws_path)
        btn_refresh = QPushButton("↻  Aggiorna")
        btn_refresh.clicked.connect(self.refresh)
        hdr.addWidget(btn_refresh)
        layout.addLayout(hdr)

        # ---- Barra filtri ----
        flt = QHBoxLayout()

        flt.addWidget(QLabel("Cerca:"))
        self.txt_search = QLineEdit()
        self.txt_search.setPlaceholderText("Codice, titolo, descrizione...")
        self.txt_search.setClearButtonEnabled(True)
        self.txt_search.textChanged.connect(self._apply_filter)
        flt.addWidget(self.txt_search, 2)

        flt.addWidget(QLabel("Checkout:"))
        self.cmb_co = QComboBox()
        self.cmb_co.addItems(["Tutti", "Mio checkout", "In checkout (altri)", "Libero"])
        self.cmb_co.currentIndexChanged.connect(self._apply_filter)
        flt.addWidget(self.cmb_co)

        layout.addLayout(flt)

        # ---- Tabella ----
        self.tbl = QTableWidget(0, 7)
        self.tbl.setHorizontalHeaderLabels([
            "File", "Codice", "Rev.", "Tipo", "Stato workflow", "Checkout", "Versione WS"
        ])
        hh = self.tbl.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        hh.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        hh.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        hh.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        hh.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        hh.setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        hh.setSectionResizeMode(6, QHeaderView.ResizeMode.ResizeToContents)
        self.tbl.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.tbl.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tbl.verticalHeader().setVisible(False)
        self.tbl.setAlternatingRowColors(True)
        self.tbl.setSortingEnabled(True)
        layout.addWidget(self.tbl)

        # ---- Pulsanti azioni ----
        btn_row = QHBoxLayout()
        self.btn_checkin   = QPushButton("📥  Check-in")
        self.btn_checkin.setObjectName("btn_primary")
        self.btn_checkin.clicked.connect(self._checkin)

        self.btn_checkout  = QPushButton("📤  Checkout")
        self.btn_checkout.clicked.connect(self._checkout)

        self.btn_undo      = QPushButton("↩  Annulla Checkout")
        self.btn_undo.setObjectName("btn_warning")
        self.btn_undo.clicked.connect(self._undo_checkout)

        self.btn_update_ws = QPushButton("⬇  Aggiorna Workspace")
        self.btn_update_ws.clicked.connect(self._update_ws)

        self.btn_delete_ws = QPushButton("🗑  Elimina file in WS")
        self.btn_delete_ws.setObjectName("btn_danger")
        self.btn_delete_ws.clicked.connect(self._delete_ws)

        for b in [self.btn_checkin, self.btn_checkout, self.btn_undo,
                  self.btn_update_ws, self.btn_delete_ws]:
            btn_row.addWidget(b)
        btn_row.addStretch()
        layout.addLayout(btn_row)

    def save_layout(self):
        """Compatibilità con MainWindow.closeEvent."""
        pass

    # ------------------------------------------------------------------
    #  Refresh – scansione fisica della cartella workspace
    # ------------------------------------------------------------------
    def refresh(self):
        if not session.is_logged_in:
            return

        user = session.user
        uid  = user["id"]
        self.lbl_user.setText(f"👤  {user['full_name']}  ({user['role']})")

        from config import load_local_config
        cfg     = load_local_config()
        ws_root = cfg.get("sw_workspace", "")
        ws_path = Path(ws_root) if ws_root else None

        if not ws_path or not ws_path.is_dir():
            self.lbl_ws_path.setText("⚠️  Workspace non configurata")
            self.tbl.setRowCount(0)
            return

        self.lbl_ws_path.setText(str(ws_path))

        # Scansiona file SolidWorks
        sw_files = []
        for pat in ("*.SLDPRT", "*.SLDASM", "*.SLDDRW",
                    "*.sldprt", "*.sldasm", "*.slddrw"):
            sw_files.extend(ws_path.glob(pat))

        # Deduplica (Windows FS case-insensitive) e ordina
        seen, unique = set(), []
        for f in sw_files:
            k = f.name.lower()
            if k not in seen:
                seen.add(k)
                unique.append(f)
        unique.sort(key=lambda f: f.name.lower())

        self.tbl.setSortingEnabled(False)
        self.tbl.setRowCount(0)
        for ws_file in unique:
            self._add_row(ws_file, uid)
        self.tbl.setSortingEnabled(True)

        self._apply_filter()

    # ------------------------------------------------------------------
    def _add_row(self, ws_file: Path, uid: int):
        code      = ws_file.stem
        ext_lower = ws_file.suffix.lower()
        doc_type  = _EXT_TO_TYPE.get(ext_lower, "")

        # Cerca nel DB preferendo: in checkout > non obsoleto > rev più alta
        if doc_type:
            doc = session.db.fetchone(
                "SELECT d.*, u.full_name AS locked_by_name "
                "FROM documents d LEFT JOIN users u ON u.id=d.locked_by "
                "WHERE d.code=? AND d.doc_type=? "
                "ORDER BY d.is_locked DESC, "
                "CASE d.state WHEN 'Obsoleto' THEN 0 ELSE 1 END DESC, "
                "d.revision DESC LIMIT 1",
                (code, doc_type),
            )
        else:
            doc = session.db.fetchone(
                "SELECT d.*, u.full_name AS locked_by_name "
                "FROM documents d LEFT JOIN users u ON u.id=d.locked_by "
                "WHERE d.code=? "
                "ORDER BY d.is_locked DESC, "
                "CASE d.state WHEN 'Obsoleto' THEN 0 ELSE 1 END DESC, "
                "d.revision DESC LIMIT 1",
                (code,),
            )

        r = self.tbl.rowCount()
        self.tbl.insertRow(r)

        # Col 0 – nome file (porta i dati per filtraggio e azioni)
        file_item = QTableWidgetItem(ws_file.name)
        file_item.setData(_ROLE_WS_PATH, str(ws_file))
        self.tbl.setItem(r, 0, file_item)

        if doc is None:
            file_item.setData(_ROLE_DOC_ID, None)
            file_item.setData(_ROLE_WS_STATUS, "unknown")
            file_item.setData(_ROLE_CO_STATE, "unknown")
            file_item.setData(_ROLE_SEARCH_TEXT, ws_file.stem.lower())
            grey = QColor("#6c7086")
            for col, txt in [
                (1, "—"), (2, "—"), (3, "—"),
                (4, "Non registrato in PDM"), (5, "—"), (6, "—")
            ]:
                it = QTableWidgetItem(txt)
                it.setForeground(grey)
                self.tbl.setItem(r, col, it)
            return

        doc_id = doc["id"]
        file_item.setData(_ROLE_DOC_ID, doc_id)

        # Testo ricercabile: codice + titolo + descrizione
        search_text = " ".join([
            doc.get("code", ""),
            doc.get("title", "") or "",
            doc.get("description", "") or "",
        ]).lower()
        file_item.setData(_ROLE_SEARCH_TEXT, search_text)

        # Col 1 – Codice
        self.tbl.setItem(r, 1, QTableWidgetItem(doc["code"]))
        # Col 2 – Rev.
        self.tbl.setItem(r, 2, QTableWidgetItem(doc["revision"]))
        # Col 3 – Tipo
        icon = TYPE_ICON.get(doc["doc_type"], "")
        self.tbl.setItem(r, 3, QTableWidgetItem(f"{icon} {doc['doc_type']}"))

        # Col 4 – Stato workflow
        state = doc["state"]
        st_item = QTableWidgetItem(state)
        st_color = WORKFLOW_STATES.get(state, {}).get("color", "#9E9E9E")
        st_item.setForeground(QColor(st_color))
        self.tbl.setItem(r, 4, st_item)

        # Col 5 – Checkout
        if doc["is_locked"]:
            locked_name = doc.get("locked_by_name") or "?"
            is_mine = (doc["locked_by"] == uid)
            if is_mine:
                co_txt   = "🔒 Mio checkout"
                co_color = "#fab387"
                co_state = "mine"
            else:
                co_txt   = f"🔒 {locked_name}"
                co_color = "#f38ba8"
                co_state = "other"
        else:
            co_txt   = "✅ Libero"
            co_color = "#a6e3a1"
            co_state = "free"
        co_item = QTableWidgetItem(co_txt)
        co_item.setForeground(QColor(co_color))
        self.tbl.setItem(r, 5, co_item)
        file_item.setData(_ROLE_CO_STATE, co_state)

        # Col 6 – Versione WS
        ws_status, vs_txt, vs_color = self._compute_ws_status(doc, ws_file, uid)
        file_item.setData(_ROLE_WS_STATUS, ws_status)
        vs_item = QTableWidgetItem(vs_txt)
        vs_item.setForeground(QColor(vs_color))
        self.tbl.setItem(r, 6, vs_item)

    # ------------------------------------------------------------------
    def _compute_ws_status(self, doc: dict, ws_file: Path, uid: int):
        """Ritorna (status_key, testo, colore)."""
        from core.checkout_manager import CheckoutManager

        has_archive = bool(doc.get("archive_path"))
        is_locked   = bool(doc["is_locked"])
        is_mine     = is_locked and (doc["locked_by"] == uid)

        if not has_archive:
            return ("not_archived", "⚠️ Non archiviato", "#f9e2af")

        sp = getattr(session, "sp", None)
        if not sp:
            return ("unknown", "? Archivio n/d", "#6c7086")

        arch_path = sp.root / doc["archive_path"]
        if not arch_path.exists():
            return ("not_archived", "⚠️ File archivio mancante", "#f38ba8")

        try:
            ws_md5   = CheckoutManager._md5(ws_file)
            arch_md5 = CheckoutManager._md5(arch_path)
        except Exception:
            return ("unknown", "? Errore confronto", "#6c7086")

        if is_mine:
            checkout_md5 = doc.get("checkout_md5") or ""
            if checkout_md5 and ws_md5 != checkout_md5:
                return ("modified", "✏️ Modificato", "#89b4fa")
            elif ws_md5 == arch_md5:
                return ("current", "🔒 In checkout", "#fab387")
            else:
                return ("modified", "✏️ Modificato", "#89b4fa")
        else:
            if ws_md5 == arch_md5:
                return ("current", "✅ Aggiornato", "#a6e3a1")
            try:
                arch_mtime = arch_path.stat().st_mtime
                ws_mtime   = ws_file.stat().st_mtime
                if arch_mtime > ws_mtime:
                    return ("outdated", "⬇ Aggiornamento disponibile", "#f9e2af")
            except Exception:
                pass
            return ("copy", "⚠️ Differisce dall'archivio", "#f38ba8")

    # ------------------------------------------------------------------
    #  Filtro (testo + stato checkout) – opera su righe già inserite
    # ------------------------------------------------------------------
    def _apply_filter(self):
        query   = self.txt_search.text().strip().lower()
        co_idx  = self.cmb_co.currentIndex()   # 0=Tutti 1=Mine 2=Altri 3=Libero
        co_map  = {1: "mine", 2: "other", 3: "free"}
        co_filter = co_map.get(co_idx)          # None → Tutti

        for row in range(self.tbl.rowCount()):
            item = self.tbl.item(row, 0)
            if item is None:
                continue

            # Filtro testo
            if query:
                search_text = item.data(_ROLE_SEARCH_TEXT) or ""
                if query not in search_text:
                    self.tbl.setRowHidden(row, True)
                    continue

            # Filtro checkout
            if co_filter:
                co_state = item.data(_ROLE_CO_STATE) or "unknown"
                if co_state != co_filter:
                    self.tbl.setRowHidden(row, True)
                    continue

            self.tbl.setRowHidden(row, False)

    # ------------------------------------------------------------------
    #  Helper selezione (considera solo righe visibili)
    # ------------------------------------------------------------------
    def _selected(self):
        """Ritorna (row, doc_id, ws_path, ws_status) o None."""
        row = self.tbl.currentRow()
        if row < 0 or self.tbl.isRowHidden(row):
            return None
        item = self.tbl.item(row, 0)
        return (
            row,
            item.data(_ROLE_DOC_ID),
            item.data(_ROLE_WS_PATH),
            item.data(_ROLE_WS_STATUS),
        )

    # ------------------------------------------------------------------
    #  Azioni
    # ------------------------------------------------------------------
    def _checkin(self):
        sel = self._selected()
        if not sel:
            QMessageBox.warning(self, "Nessuna selezione", "Selezionare un file dalla lista.")
            return
        _, doc_id, _, _ = sel
        if not doc_id:
            QMessageBox.warning(self, "Non registrato",
                                "Il file non è registrato nel PDM.")
            return
        from ui.checkin_dialog import CheckinDialog
        dlg = CheckinDialog(doc_id, parent=self)
        if dlg.exec() == CheckinDialog.DialogCode.Accepted:
            self.refresh()

    def _checkout(self):
        sel = self._selected()
        if not sel:
            QMessageBox.warning(self, "Nessuna selezione", "Selezionare un file dalla lista.")
            return
        _, doc_id, _, ws_status = sel
        if not doc_id:
            QMessageBox.warning(self, "Non registrato",
                                "Il file non è registrato nel PDM.")
            return
        if ws_status == "not_archived":
            QMessageBox.warning(
                self, "Non archiviato",
                "Il file non è ancora in archivio PDM.\n"
                "Usare 'Importa in PDM' dalla scheda documento prima del checkout."
            )
            return
        try:
            session.checkout.checkout(doc_id)
            self.refresh()
        except Exception as e:
            QMessageBox.critical(self, "Errore checkout", str(e))

    def _undo_checkout(self):
        sel = self._selected()
        if not sel:
            QMessageBox.warning(self, "Nessuna selezione", "Selezionare un file dalla lista.")
            return
        _, doc_id, _, _ = sel
        if not doc_id:
            return
        r = QMessageBox.question(
            self, "Annulla checkout",
            "Annullare il checkout senza archiviare?\n"
            "Le modifiche locali al file non verranno salvate in PDM.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if r == QMessageBox.StandardButton.Yes:
            try:
                session.checkout.undo_checkout(doc_id, delete_from_workspace=False)
                self.refresh()
            except Exception as e:
                QMessageBox.critical(self, "Errore", str(e))

    def _update_ws(self):
        """Copia il file dall'archivio PDM alla workspace (sovrascrive)."""
        sel = self._selected()
        if not sel:
            QMessageBox.warning(self, "Nessuna selezione", "Selezionare un file dalla lista.")
            return
        _, doc_id, _, ws_status = sel
        if not doc_id:
            return
        if ws_status == "not_archived":
            QMessageBox.warning(self, "Non archiviato",
                                "Nessun file in archivio da cui aggiornare.")
            return
        if ws_status == "modified":
            r = QMessageBox.question(
                self, "File modificato localmente",
                "Il file è stato modificato nella workspace.\n"
                "Sovrascrivere con la versione dell'archivio?\n\n"
                "Le modifiche locali andranno perse.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if r != QMessageBox.StandardButton.Yes:
                return
        try:
            from config import load_local_config
            ws = Path(load_local_config().get("sw_workspace", ""))
            exported = session.files.export_to_workspace(doc_id, ws)
            self.refresh()
            if exported:
                QMessageBox.information(
                    self, "Workspace aggiornata",
                    f"File copiato dall'archivio:\n{exported[0]}"
                )
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

    def _delete_ws(self):
        """Elimina fisicamente il file dalla cartella workspace."""
        sel = self._selected()
        if not sel:
            QMessageBox.warning(self, "Nessuna selezione", "Selezionare un file dalla lista.")
            return
        _, doc_id, ws_path, _ = sel
        if not ws_path:
            return

        # Blocca se il file è in checkout dell'utente corrente
        if doc_id:
            doc = session.db.fetchone(
                "SELECT is_locked, locked_by FROM documents WHERE id=?", (doc_id,)
            )
            if doc and doc["is_locked"] and doc["locked_by"] == session.user["id"]:
                QMessageBox.warning(
                    self, "File in checkout",
                    "Il file è in checkout. Prima fare Check-in o Annulla Checkout."
                )
                return

        r = QMessageBox.question(
            self, "Elimina dalla workspace",
            f"Eliminare il file:\n{ws_path}\n\n"
            "Il file nell'archivio PDM non verrà modificato.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if r == QMessageBox.StandardButton.Yes:
            try:
                p = Path(ws_path)
                from core.checkout_manager import CheckoutManager
                CheckoutManager._set_writable(p)
                p.unlink()
                if doc_id:
                    try:
                        session.checkout.remove_from_workspace(doc_id, delete_file=False)
                    except Exception:
                        pass
                self.refresh()
            except Exception as e:
                QMessageBox.critical(self, "Errore", str(e))

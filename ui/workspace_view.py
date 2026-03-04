# =============================================================================
#  ui/workspace_view.py  –  Vista workspace (checkout + copie)  v2.0
# =============================================================================
from pathlib import Path
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
    QMessageBox, QGroupBox, QSplitter
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor

from config import WORKFLOW_STATES
from core.checkout_manager import READONLY_STATES
from ui.session import session
from ui.styles import TYPE_ICON


class WorkspaceView(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._build_ui()

    # ------------------------------------------------------------------
    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(8)
        layout.setContentsMargins(8, 8, 8, 8)

        # Header
        hdr = QHBoxLayout()
        self.lbl_user = QLabel("\u2014")
        self.lbl_user.setObjectName("title_label")
        hdr.addWidget(self.lbl_user)
        hdr.addStretch()
        btn_refresh = QPushButton("\u21bb  Aggiorna")
        btn_refresh.clicked.connect(self.refresh)
        hdr.addWidget(btn_refresh)
        layout.addLayout(hdr)

        splitter = QSplitter(Qt.Orientation.Vertical)

        # ---- Sezione 1: Documenti in checkout (modificabili) ----
        grp_co = QGroupBox("\U0001f512  Documenti in checkout (modificabili)")
        co_layout = QVBoxLayout(grp_co)

        self.tbl_checkout = QTableWidget(0, 6)
        self.tbl_checkout.setHorizontalHeaderLabels([
            "Codice", "Rev.", "Tipo", "Titolo", "Stato", "Data checkout"
        ])
        self.tbl_checkout.horizontalHeader().setSectionResizeMode(
            3, QHeaderView.ResizeMode.Stretch
        )
        self.tbl_checkout.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows
        )
        self.tbl_checkout.setEditTriggers(
            QTableWidget.EditTrigger.NoEditTriggers
        )
        self.tbl_checkout.verticalHeader().setVisible(False)
        co_layout.addWidget(self.tbl_checkout)

        btn_row1 = QHBoxLayout()
        btn_checkin = QPushButton("\U0001f4e5  Check-in")
        btn_checkin.setObjectName("btn_primary")
        btn_checkin.clicked.connect(self._checkin_selected)
        btn_undo = QPushButton("\u21a9\ufe0f  Annulla checkout")
        btn_undo.setObjectName("btn_warning")
        btn_undo.clicked.connect(self._undo_checkout_selected)
        btn_open = QPushButton("\U0001f4c2  Apri in workspace")
        btn_open.clicked.connect(self._open_checkout_file)
        btn_row1.addWidget(btn_checkin)
        btn_row1.addWidget(btn_undo)
        btn_row1.addWidget(btn_open)
        btn_row1.addStretch()
        co_layout.addLayout(btn_row1)
        splitter.addWidget(grp_co)

        # ---- Sezione 2: File in workspace (copie senza lock) ----
        grp_ws = QGroupBox("\U0001f4c1  File in workspace (copie / consultazione)")
        ws_layout = QVBoxLayout(grp_ws)

        self.tbl_workspace = QTableWidget(0, 5)
        self.tbl_workspace.setHorizontalHeaderLabels([
            "Codice", "Tipo", "Ruolo", "Stato", "Data copia"
        ])
        self.tbl_workspace.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeMode.Stretch
        )
        self.tbl_workspace.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows
        )
        self.tbl_workspace.setEditTriggers(
            QTableWidget.EditTrigger.NoEditTriggers
        )
        self.tbl_workspace.verticalHeader().setVisible(False)
        ws_layout.addWidget(self.tbl_workspace)

        btn_row2 = QHBoxLayout()
        btn_remove = QPushButton("\U0001f5d1  Rimuovi dalla workspace")
        btn_remove.setObjectName("btn_danger")
        btn_remove.clicked.connect(self._remove_selected)
        btn_open_ws = QPushButton("\U0001f4c2  Apri file")
        btn_open_ws.clicked.connect(self._open_ws_file)
        btn_row2.addWidget(btn_remove)
        btn_row2.addWidget(btn_open_ws)
        btn_row2.addStretch()
        ws_layout.addLayout(btn_row2)
        splitter.addWidget(grp_ws)

        layout.addWidget(splitter)

    # ------------------------------------------------------------------
    #  Refresh
    # ------------------------------------------------------------------
    def refresh(self):
        if not session.is_logged_in:
            return
        user = session.user
        uid = user["id"]
        self.lbl_user.setText(
            f"\U0001f464  {user['full_name']}  ({user['role']})"
        )

        # ---- Checkout (file con lock mio) ----
        checkout_docs = session.checkout.get_checked_out_by_user(uid)
        self.tbl_checkout.setRowCount(0)
        for doc in checkout_docs:
            r = self.tbl_checkout.rowCount()
            self.tbl_checkout.insertRow(r)
            self.tbl_checkout.setItem(r, 0, QTableWidgetItem(doc["code"]))
            self.tbl_checkout.setItem(r, 1, QTableWidgetItem(doc["revision"]))
            icon = TYPE_ICON.get(doc["doc_type"], "")
            self.tbl_checkout.setItem(
                r, 2, QTableWidgetItem(f"{icon} {doc['doc_type']}")
            )
            self.tbl_checkout.setItem(r, 3, QTableWidgetItem(doc.get("title") or ""))

            # Stato con colore
            st = doc["state"]
            state_item = QTableWidgetItem(st)
            color = WORKFLOW_STATES.get(st, {}).get("color", "#9E9E9E")
            state_item.setForeground(QColor(color))
            self.tbl_checkout.setItem(r, 4, state_item)

            locked_at = doc.get("locked_at", "")
            self.tbl_checkout.setItem(
                r, 5, QTableWidgetItem(locked_at[:19] if locked_at else "")
            )
            self.tbl_checkout.item(r, 0).setData(
                Qt.ItemDataRole.UserRole, doc["id"]
            )

        # ---- Workspace files (copie senza lock) ----
        ws_files = session.checkout.get_workspace_files(uid)
        # Filtra via i file in checkout (gia mostrati sopra)
        copies = [
            wf for wf in ws_files
            if wf.get("role") in ("component", "consultation")
        ]
        self.tbl_workspace.setRowCount(0)
        for wf in copies:
            r = self.tbl_workspace.rowCount()
            self.tbl_workspace.insertRow(r)
            self.tbl_workspace.setItem(r, 0, QTableWidgetItem(wf["code"]))
            icon = TYPE_ICON.get(wf.get("doc_type", ""), "")
            self.tbl_workspace.setItem(
                r, 1, QTableWidgetItem(f"{icon} {wf.get('doc_type', '')}")
            )
            role_label = {
                "component": "Componente (copia)",
                "consultation": "Consultazione",
            }.get(wf.get("role", ""), wf.get("role", ""))
            self.tbl_workspace.setItem(r, 2, QTableWidgetItem(role_label))

            st = wf.get("state", "")
            state_item = QTableWidgetItem(st)
            color = WORKFLOW_STATES.get(st, {}).get("color", "#9E9E9E")
            state_item.setForeground(QColor(color))
            self.tbl_workspace.setItem(r, 3, state_item)

            copied_at = wf.get("copied_at", "")
            self.tbl_workspace.setItem(
                r, 4, QTableWidgetItem(str(copied_at)[:19] if copied_at else "")
            )
            self.tbl_workspace.item(r, 0).setData(
                Qt.ItemDataRole.UserRole, wf["document_id"]
            )

    # ------------------------------------------------------------------
    #  Azioni checkout
    # ------------------------------------------------------------------
    def _checkin_selected(self):
        row = self.tbl_checkout.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Nessuna selezione",
                                "Selezionare un documento in checkout.")
            return
        doc_id = self.tbl_checkout.item(row, 0).data(Qt.ItemDataRole.UserRole)
        from ui.checkin_dialog import CheckinDialog
        dlg = CheckinDialog(doc_id, parent=self)
        if dlg.exec() == CheckinDialog.DialogCode.Accepted:
            self.refresh()

    def _undo_checkout_selected(self):
        row = self.tbl_checkout.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Nessuna selezione",
                                "Selezionare un documento in checkout.")
            return
        doc_id = self.tbl_checkout.item(row, 0).data(Qt.ItemDataRole.UserRole)
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

    def _open_checkout_file(self):
        row = self.tbl_checkout.currentRow()
        if row < 0:
            return
        doc_id = self.tbl_checkout.item(row, 0).data(Qt.ItemDataRole.UserRole)
        try:
            session.files.open_from_workspace(doc_id)
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

    # ------------------------------------------------------------------
    #  Azioni workspace (copie)
    # ------------------------------------------------------------------
    def _remove_selected(self):
        row = self.tbl_workspace.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Nessuna selezione",
                                "Selezionare un file dalla lista workspace.")
            return
        doc_id = self.tbl_workspace.item(row, 0).data(Qt.ItemDataRole.UserRole)
        r = QMessageBox.question(
            self, "Rimuovi dalla workspace",
            "Rimuovere il file dalla workspace?\n"
            "Il file in archivio non verra modificato.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if r == QMessageBox.StandardButton.Yes:
            try:
                session.checkout.remove_from_workspace(doc_id, delete_file=True)
                self.refresh()
            except Exception as e:
                QMessageBox.critical(self, "Errore", str(e))

    def _open_ws_file(self):
        row = self.tbl_workspace.currentRow()
        if row < 0:
            return
        doc_id = self.tbl_workspace.item(row, 0).data(Qt.ItemDataRole.UserRole)
        try:
            session.files.open_from_workspace(doc_id)
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

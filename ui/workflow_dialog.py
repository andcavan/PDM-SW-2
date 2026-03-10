# =============================================================================
#  ui/workflow_dialog.py  –  Cambio stato workflow documento  v3.0
# =============================================================================
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QComboBox, QTextEdit, QGroupBox,
    QFormLayout, QMessageBox
)
from PyQt6.QtCore import Qt

from config import WORKFLOW_STATES
from ui.session import session
from ui.styles import STATE_BADGE_STYLE


class WorkflowDialog(QDialog):
    def __init__(self, document_id: int, parent=None, skip_r2: bool = False):
        super().__init__(parent)
        self.document_id = document_id
        self.skip_r2 = skip_r2
        self.setWindowTitle("Workflow")
        self.setMinimumWidth(420)
        self._build_ui()
        self._load()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(20, 20, 20, 20)

        # Stato corrente
        grp_curr = QGroupBox("Stato corrente")
        form_curr = QFormLayout(grp_curr)
        self.lbl_state = QLabel("—")
        form_curr.addRow("Stato:", self.lbl_state)
        self.lbl_code = QLabel("—")
        form_curr.addRow("Documento:", self.lbl_code)
        self.lbl_lock = QLabel("")
        form_curr.addRow("Checkout:", self.lbl_lock)
        layout.addWidget(grp_curr)

        # Cambio stato
        grp_ch = QGroupBox("Cambia stato")
        form_ch = QFormLayout(grp_ch)
        self.cmb_target  = QComboBox()
        self.lbl_dot     = QLabel()
        form_ch.addRow("Nuovo stato:", self.cmb_target)
        form_ch.addRow("", self.lbl_dot)
        self.txt_notes = QTextEdit()
        self.txt_notes.setMaximumHeight(70)
        self.txt_notes.setPlaceholderText("Note (opzionale)…")
        form_ch.addRow("Note:", self.txt_notes)
        layout.addWidget(grp_ch)

        # Bottoni
        btn_row = QHBoxLayout()
        btn_row.addStretch()

        btn_close = QPushButton("Chiudi")
        btn_close.clicked.connect(self.accept)
        btn_row.addWidget(btn_close)

        btn_apply = QPushButton("Applica transizione")
        btn_apply.setObjectName("btn_primary")
        btn_apply.clicked.connect(self._apply_transition)
        self._btn_apply = btn_apply
        btn_row.addWidget(btn_apply)
        layout.addLayout(btn_row)

    def _load(self):
        doc = session.files.get_document(self.document_id)
        if not doc:
            return

        state  = doc["state"]
        badge  = STATE_BADGE_STYLE.get(state, "")
        self.lbl_state.setText(f"<span style='{badge}'>{state}</span>")
        self.lbl_state.setTextFormat(Qt.TextFormat.RichText)
        self.lbl_code.setText(f"{doc['code']} rev.{doc['revision']}  –  {doc['title']}")

        # Mostra stato checkout
        if doc.get("is_locked"):
            locker = doc.get("locked_by_name", f"utente {doc.get('locked_by','')}")
            self.lbl_lock.setText(f"⚠️  In checkout: {locker}")
            self.lbl_lock.setStyleSheet("color:#f38ba8; font-weight:bold;")
        else:
            self.lbl_lock.setText("—")
            self.lbl_lock.setStyleSheet("")

        # Blocca transizioni se in checkout
        locked = bool(doc.get("is_locked"))
        self._btn_apply.setEnabled(not locked)

        transitions = session.workflow.get_available_transitions(state)
        self.cmb_target.blockSignals(True)
        self.cmb_target.clear()
        for t in transitions:
            self.cmb_target.addItem(t)

        if not transitions:
            self.cmb_target.addItem("Nessuna transizione disponibile")
            self.cmb_target.setEnabled(False)
            self._btn_apply.setEnabled(False)
        else:
            self.cmb_target.setEnabled(not locked)

        self.cmb_target.blockSignals(False)
        self.cmb_target.currentTextChanged.connect(self._update_dot)
        self._update_dot()

    def _update_dot(self):
        t     = self.cmb_target.currentText()
        color = WORKFLOW_STATES.get(t, {}).get("color", "#9E9E9E")
        self.lbl_dot.setText(
            f"<span style='background:{color};color:white;"
            f"border-radius:4px;padding:2px 10px;'>{t}</span>"
        )
        self.lbl_dot.setTextFormat(Qt.TextFormat.RichText)

    def _apply_transition(self):
        target = self.cmb_target.currentText()
        if not target or "Nessuna" in target:
            return
        notes = self.txt_notes.toPlainText().strip()
        try:
            session.workflow.change_state(
                self.document_id, target, session.user["id"], notes,
                shared_paths=session.sp,
                skip_r2=self.skip_r2,
            )
            QMessageBox.information(
                self, "OK", f"Stato cambiato in: {target}"
            )
            self._load()
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

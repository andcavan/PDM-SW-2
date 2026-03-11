# =============================================================================
#  ui/asm_release_check_dialog.py  –  Verifica e rilascio componenti BOM
#
#  Mostrata prima del rilascio di un Assieme.
#  Regole:
#   - Tutti i componenti non ancora in "Rilasciato" vengono elencati.
#   - I componenti NON CODIFICATI (description contiene "non codificato")
#     sono mostrati in grigio a titolo informativo e NON bloccano il rilascio.
#   - Tra i componenti CODIFICATI non rilasciati:
#       * Se ANCHE UNO non è rilasciabile → blocco rigido (pulsante disabilitato).
#       * Se TUTTI sono rilasciabili → "Rilascia tutti e procedi" li rilascia
#         in ordine bottom-up e chiude con Accept.
#   - I componenti già in stato Rilasciato non subiscono alcun cambiamento.
# =============================================================================
from __future__ import annotations

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QTableWidget, QTableWidgetItem,
    QHeaderView, QMessageBox,
    QAbstractItemView,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor

from ui.session import session
from ui.styles import STATE_BADGE_STYLE, TYPE_ICON


def _is_non_coded(comp: dict) -> bool:
    """Componente non generato con 'Crea codice': machine_id è NULL."""
    return comp.get("machine_id") is None


class AsmReleaseCheckDialog(QDialog):
    """
    Dialog che elenca i componenti BOM di un assieme non in stato "Rilasciato".

    Solo i componenti CODIFICATI partecipano al controllo di rilasciabilità.
    Il pulsante "Rilascia tutti e procedi" è abilitato solo se tutti i componenti
    codificati non rilasciati sono rilasciabili. I non-codificati sono mostrati
    a titolo informativo e non bloccano il flusso.
    """

    _COL_CODE  = 0
    _COL_REV   = 1
    _COL_TYPE  = 2
    _COL_TITLE = 3
    _COL_STATE = 4
    _COL_CAN   = 5
    _NCOLS     = 6

    def __init__(self, asm_doc_id: int, parent=None):
        super().__init__(parent)
        self.asm_doc_id = asm_doc_id

        # BOM piatta DFS pre-order; invertita = bottom-up per il rilascio
        self._flat = session.asm.get_bom_flat(asm_doc_id)
        self._unreleased = [c for c in self._flat if c["state"] != "Rilasciato"]

        self.setWindowTitle("Rilascio assieme — verifica componenti")
        self.setMinimumSize(760, 380)
        self._build_ui()
        self._populate()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        n = len(self._unreleased)
        if n == 0:
            msg = "<b>Tutti i componenti sono già in stato Rilasciato.</b>"
        else:
            msg = (
                f"<b>{n} componente{'i' if n != 1 else ''} "
                f"non rilasciato{'i' if n != 1 else ''}.</b>"
                "  Verificare lo stato prima di procedere."
            )
        lbl = QLabel(msg)
        lbl.setWordWrap(True)
        layout.addWidget(lbl)

        # Tabella
        self.tbl = QTableWidget(0, self._NCOLS)
        self.tbl.setHorizontalHeaderLabels(
            ["Codice", "Rev", "Tipo", "Titolo", "Stato", "Rilasciabile"]
        )
        self.tbl.horizontalHeader().setSectionResizeMode(
            self._COL_TITLE, QHeaderView.ResizeMode.Stretch
        )
        self.tbl.setColumnWidth(self._COL_CODE,  120)
        self.tbl.setColumnWidth(self._COL_REV,    40)
        self.tbl.setColumnWidth(self._COL_TYPE,   80)
        self.tbl.setColumnWidth(self._COL_STATE, 110)
        self.tbl.setColumnWidth(self._COL_CAN,    90)
        self.tbl.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tbl.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        self.tbl.verticalHeader().setVisible(False)
        layout.addWidget(self.tbl)

        # Label stato globale
        self.lbl_status = QLabel()
        self.lbl_status.setWordWrap(True)
        layout.addWidget(self.lbl_status)

        # Bottoni
        btn_row = QHBoxLayout()
        btn_row.addStretch()

        btn_cancel = QPushButton("Annulla")
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_cancel)

        self.btn_release = QPushButton("Rilascia tutti e procedi")
        self.btn_release.setObjectName("btn_primary")
        self.btn_release.clicked.connect(self._do_release)
        btn_row.addWidget(self.btn_release)

        layout.addLayout(btn_row)

    def _populate(self):
        """Riempie la tabella e aggiorna lo stato del pulsante."""
        self.tbl.setRowCount(0)

        if not self._unreleased:
            self.lbl_status.setText(
                "<span style='color:#2e7d32;font-weight:bold;'>"
                "✅ Tutti i componenti sono già Rilasciati."
                "</span>"
            )
            self.btn_release.setText("Procedi")
            return

        # Analisi componenti codificati
        blocked_coded = 0
        self._coded_unreleased: list[dict] = []

        for comp in self._unreleased:
            non_coded = _is_non_coded(comp)
            if non_coded:
                can_release = False
                reason = ""
            else:
                transitions = session.workflow.get_available_transitions(comp["state"])
                can_release = ("Rilasciato" in transitions) and bool(comp.get("archive_path"))
                if not can_release:
                    blocked_coded += 1
                    reason = (
                        "File non archiviato (check-in non eseguito)"
                        if not comp.get("archive_path")
                        else "Transizione a Rilasciato non disponibile dallo stato corrente"
                    )
                else:
                    reason = ""
                self._coded_unreleased.append(comp)

            row = self.tbl.rowCount()
            self.tbl.insertRow(row)

            tipo = comp.get("doc_type", "")
            icon = TYPE_ICON.get(tipo, "")
            self._set_item(row, self._COL_CODE,  comp.get("code", ""))
            self._set_item(row, self._COL_REV,   comp.get("revision", ""))
            self._set_item(row, self._COL_TYPE,  f"{icon} {tipo}")
            self._set_item(row, self._COL_TITLE, comp.get("title") or "")

            # Stato con badge colore
            state = comp.get("state", "")
            state_item = QTableWidgetItem(state)
            badge_color = {
                "In Lavorazione": "#2196F3",
                "In Revisione":   "#FF9800",
                "Obsoleto":       "#757575",
            }.get(state, "#9E9E9E")
            state_item.setForeground(QColor("white"))
            state_item.setBackground(QColor(badge_color))
            state_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.tbl.setItem(row, self._COL_STATE, state_item)

            # Colonna rilasciabile
            if non_coded:
                can_lbl = "—"
            else:
                can_lbl = "✅ Sì" if can_release else "⛔ No"
            can_item = QTableWidgetItem(can_lbl)
            can_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            if not non_coded and not can_release and reason:
                can_item.setToolTip(reason)
            self.tbl.setItem(row, self._COL_CAN, can_item)

            # Sfondo: grigio per non-codificati, rosso pallido per codificati bloccati
            if non_coded:
                bg = QColor("#e0e0e0")
                fg = QColor("#757575")
                for col in range(self._NCOLS):
                    item = self.tbl.item(row, col)
                    if item:
                        item.setBackground(bg)
                        item.setForeground(fg)
            elif not can_release:
                for col in range(self._NCOLS):
                    item = self.tbl.item(row, col)
                    if item:
                        item.setBackground(QColor("#fde8e8"))
                        item.setForeground(QColor("#333333"))

        # Label e pulsante in base ai soli codificati
        if blocked_coded > 0:
            self.lbl_status.setText(
                f"<span style='color:#c62828;font-weight:bold;'>"
                f"⛔ {blocked_coded} componente{'i' if blocked_coded != 1 else ''} codificato"
                f"{'i' if blocked_coded != 1 else ''} non "
                f"{'possono' if blocked_coded != 1 else 'può'} essere rilasciato"
                f"{'i' if blocked_coded != 1 else ''} — risolvere prima di procedere."
                f"</span>"
            )
            self.btn_release.setEnabled(False)
        else:
            coded_count = len(self._coded_unreleased)
            if coded_count > 0:
                self.lbl_status.setText(
                    "<span style='color:#2e7d32;font-weight:bold;'>"
                    "✅ Tutti i componenti codificati possono essere rilasciati."
                    "</span>"
                )
            else:
                self.lbl_status.setText(
                    "<span style='color:#2e7d32;font-weight:bold;'>"
                    "✅ Nessun componente codificato da rilasciare."
                    "</span>"
                )
            self.btn_release.setEnabled(True)

    def _set_item(self, row: int, col: int, text: str):
        item = QTableWidgetItem(text)
        self.tbl.setItem(row, col, item)

    def _do_release(self):
        """
        Rilascia tutti i componenti CODIFICATI non ancora in Rilasciato,
        in ordine bottom-up. I non-codificati e quelli già Rilasciati
        non vengono toccati.
        """
        if not self._coded_unreleased:
            self.accept()
            return

        coded_ids = {c["child_id"] for c in self._coded_unreleased}
        to_release = [
            c for c in reversed(self._flat)
            if c["child_id"] in coded_ids
        ]

        errors: list[str] = []
        for comp in to_release:
            try:
                session.workflow.change_state(
                    comp["child_id"], "Rilasciato",
                    session.user["id"],
                    notes="Rilascio contestuale assieme",
                    shared_paths=session.sp,
                )
            except Exception as e:
                errors.append(f"{comp['code']} rev.{comp['revision']}: {e}")

        if errors:
            QMessageBox.warning(
                self, "Errori durante il rilascio",
                "Alcuni componenti non sono stati rilasciati:\n\n"
                + "\n".join(f"• {e}" for e in errors)
            )

        self.accept()

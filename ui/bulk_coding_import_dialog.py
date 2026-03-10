# =============================================================================
#  ui/bulk_coding_import_dialog.py  –  Generazione / importazione massiva
#  Macchine e Gruppi coerenti con lo schema di codifica configurato
# =============================================================================
from __future__ import annotations
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QGroupBox, QFormLayout, QTabWidget, QWidget,
    QTableWidget, QTableWidgetItem, QHeaderView, QTextEdit,
    QMessageBox, QComboBox
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor

from ui.session import session


# ---------------------------------------------------------------------------
# Utilità: generatori di serie
# ---------------------------------------------------------------------------

def _alpha_to_int(code: str) -> int:
    n = 0
    for c in code.upper():
        n = n * 26 + (ord(c) - ord('A'))
    return n


def _int_to_alpha(n: int, length: int) -> str:
    chars = []
    for _ in range(length):
        chars.append(chr(n % 26 + ord('A')))
        n //= 26
    return ''.join(reversed(chars))


def generate_series(first: str, last: str, code_type: str) -> list[str]:
    """Genera tutti i codici tra first e last inclusi (ALPHA o NUM)."""
    first, last = first.upper(), last.upper()
    length = len(first)
    if len(last) != length:
        return []
    if code_type == "ALPHA":
        s = _alpha_to_int(first)
        e = _alpha_to_int(last)
        if s > e:
            s, e = e, s
        return [_int_to_alpha(i, length) for i in range(s, e + 1)]
    else:  # NUM
        try:
            s, e = int(first), int(last)
        except ValueError:
            return []
        if s > e:
            s, e = e, s
        return [str(i).zfill(length) for i in range(s, e + 1)]


# ---------------------------------------------------------------------------
# Dialog principale
# ---------------------------------------------------------------------------

class BulkCodingImportDialog(QDialog):
    """
    Wizard per generare o importare una serie di Macchine o Gruppi.

    mode  = "machines"  → gestisce tabella machines
    mode  = "groups"    → gestisce tabella machine_groups per machine_id
    """

    _STATUS_NEW   = "✅ Nuovo"
    _STATUS_EXIST = "ℹ️ Esistente"
    _STATUS_ERR   = "⚠️ Errore"

    def __init__(self, mode: str = "machines",
                 machine_id: int | None = None,
                 machine_code: str = "",
                 parent=None):
        super().__init__(parent)
        self._mode        = mode          # "machines" | "groups"
        self._machine_id  = machine_id
        self._machine_code = machine_code
        self._preview_rows: list[dict] = []   # [{"code","desc","status","note"}]

        cfg = session.coding.get_scheme_config()
        if mode == "machines":
            self._code_type   = cfg.mach_code_type
            self._code_length = cfg.mach_code_length
            title = "Genera / Importa Macchine"
        else:
            self._code_type   = cfg.grp_code_type
            self._code_length = cfg.grp_code_length
            title = f"Genera / Importa Gruppi  —  Macchina: {machine_code}"

        self.setWindowTitle(title)
        self.setMinimumSize(700, 580)
        self._build_ui()

    # ==================================================================
    # UI
    # ==================================================================
    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setSpacing(10)
        root.setContentsMargins(16, 16, 16, 16)

        # Info schema
        cfg = session.coding.get_scheme_config()
        what = "Macchina" if self._mode == "machines" else "Gruppo"
        info = QLabel(
            f"Schema attivo — {what}: <b>{self._code_type}</b>, "
            f"<b>{self._code_length}</b> caratteri"
        )
        root.addWidget(info)

        # Tab input
        self._tabs = QTabWidget()
        self._tabs.addTab(self._build_serie_tab(), "Genera serie")
        self._tabs.addTab(self._build_lista_tab(), "Lista libera")
        root.addWidget(self._tabs)

        # Pulsante genera/analizza
        btn_gen = QPushButton("Genera / Analizza →")
        btn_gen.setObjectName("btn_primary")
        btn_gen.clicked.connect(self._run_analysis)
        root.addWidget(btn_gen, alignment=Qt.AlignmentFlag.AlignRight)

        # Preview tabella
        prev_grp = QGroupBox("Anteprima")
        prev_lyt = QVBoxLayout(prev_grp)
        self._tbl = QTableWidget(0, 4)
        self._tbl.setHorizontalHeaderLabels(["Stato", "Codice", "Descrizione", "Note"])
        self._tbl.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self._tbl.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self._tbl.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self._tbl.verticalHeader().setVisible(False)
        prev_lyt.addWidget(self._tbl)
        self._lbl_summary = QLabel("")
        self._lbl_summary.setObjectName("subtitle_label")
        prev_lyt.addWidget(self._lbl_summary)
        root.addWidget(prev_grp, stretch=1)

        # Pulsanti finali
        btn_row = QHBoxLayout()
        btn_close = QPushButton("Chiudi")
        btn_close.clicked.connect(self.reject)
        self._btn_import = QPushButton("Importa validi (0)")
        self._btn_import.setObjectName("btn_primary")
        self._btn_import.setEnabled(False)
        self._btn_import.clicked.connect(self._do_import)
        btn_row.addWidget(btn_close)
        btn_row.addStretch()
        btn_row.addWidget(self._btn_import)
        root.addLayout(btn_row)

    # ------------------------------------------------------------------
    def _build_serie_tab(self) -> QWidget:
        w = QWidget()
        form = QFormLayout(w)
        form.setSpacing(10)
        form.setContentsMargins(12, 12, 12, 12)

        if self._code_type == "ALPHA":
            placeholder_first = "A" * self._code_length
            placeholder_last  = "Z" * self._code_length
            hint = f"Lettere A-Z, {self._code_length} caratteri (es. {placeholder_first} → {placeholder_last})"
        else:
            placeholder_first = "0" * self._code_length
            placeholder_last  = "9" * self._code_length
            hint = f"Cifre 0-9, {self._code_length} caratteri (es. {placeholder_first} → {placeholder_last})"

        form.addRow(QLabel(f"<i>{hint}</i>"))

        self._first = QLineEdit()
        self._first.setMaximumWidth(120)
        self._first.setPlaceholderText(placeholder_first)
        self._first.textChanged.connect(lambda t: self._first.setText(t.upper()))
        form.addRow("Primo codice:", self._first)

        self._last = QLineEdit()
        self._last.setMaximumWidth(120)
        self._last.setPlaceholderText(placeholder_last)
        self._last.textChanged.connect(lambda t: self._last.setText(t.upper()))
        form.addRow("Ultimo codice:", self._last)

        self._desc_tmpl = QLineEdit("{CODE}")
        self._desc_tmpl.setPlaceholderText("Descrizione auto (usa {CODE})")
        form.addRow("Template descrizione:", self._desc_tmpl)

        lbl_note = QLabel(
            "<i>{CODE} viene sostituito col codice generato. "
            "Lascia vuoto per descrizione vuota.</i>"
        )
        lbl_note.setWordWrap(True)
        form.addRow("", lbl_note)
        return w

    # ------------------------------------------------------------------
    def _build_lista_tab(self) -> QWidget:
        w = QWidget()
        lyt = QVBoxLayout(w)
        lyt.setContentsMargins(12, 12, 12, 12)

        lbl = QLabel(
            "Inserisci una riga per codice nel formato  <b>CODICE ; Descrizione</b>\n"
            "(la descrizione è opzionale)"
        )
        lbl.setWordWrap(True)
        lyt.addWidget(lbl)

        self._txt_lista = QTextEdit()
        self._txt_lista.setPlaceholderText(
            "ABC ; Macchina ABC\n"
            "DEF ; Macchina DEF\n"
            "GHI"
        )
        lyt.addWidget(self._txt_lista)
        return w

    # ==================================================================
    # Logica analisi
    # ==================================================================
    def _run_analysis(self):
        """Legge input dall'tab attivo, valida e popola la preview."""
        if self._tabs.currentIndex() == 0:
            rows = self._rows_from_serie()
        else:
            rows = self._rows_from_lista()

        if not rows:
            QMessageBox.warning(self, "Nessun dato", "Nessun codice da analizzare.")
            return

        # Carica codici esistenti per confronto rapido
        if self._mode == "machines":
            existing = {m["code"] for m in session.coding.get_machines(only_active=False)}
        else:
            existing = {g["code"] for g in session.coding.get_groups(
                self._machine_id, only_active=False
            )} if self._machine_id else set()

        preview = []
        for code, desc in rows:
            ok, err = session.coding.validate_code_string(
                code, self._code_type, self._code_length
            )
            if not ok:
                preview.append({"code": code, "desc": desc,
                                 "status": self._STATUS_ERR, "note": err})
            elif code in existing:
                preview.append({"code": code, "desc": desc,
                                 "status": self._STATUS_EXIST, "note": "già presente"})
            else:
                preview.append({"code": code, "desc": desc,
                                 "status": self._STATUS_NEW, "note": ""})

        self._preview_rows = preview
        self._refresh_preview()

    def _rows_from_serie(self) -> list[tuple[str, str]]:
        first = self._first.text().strip().upper()
        last  = self._last.text().strip().upper()
        tmpl  = self._desc_tmpl.text().strip()

        if not first or not last:
            return []

        codes = generate_series(first, last, self._code_type)
        if len(codes) > 500:
            r = QMessageBox.question(
                self, "Serie molto grande",
                f"La serie contiene {len(codes)} codici. Continuare?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if r != QMessageBox.StandardButton.Yes:
                return []

        return [(c, tmpl.replace("{CODE}", c) if tmpl else "") for c in codes]

    def _rows_from_lista(self) -> list[tuple[str, str]]:
        rows = []
        for line in self._txt_lista.toPlainText().splitlines():
            line = line.strip()
            if not line:
                continue
            parts = line.split(";", 1)
            code = parts[0].strip().upper()
            desc = parts[1].strip() if len(parts) > 1 else ""
            rows.append((code, desc))
        return rows

    # ==================================================================
    # Preview
    # ==================================================================
    def _refresh_preview(self):
        self._tbl.setRowCount(0)
        n_new = sum(1 for r in self._preview_rows if r["status"] == self._STATUS_NEW)
        n_ex  = sum(1 for r in self._preview_rows if r["status"] == self._STATUS_EXIST)
        n_err = sum(1 for r in self._preview_rows if r["status"] == self._STATUS_ERR)

        for r in self._preview_rows:
            row = self._tbl.rowCount()
            self._tbl.insertRow(row)
            for col, val in enumerate([r["status"], r["code"], r["desc"], r["note"]]):
                item = QTableWidgetItem(val)
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                if r["status"] == self._STATUS_NEW:
                    item.setForeground(QColor("#a6e3a1"))
                elif r["status"] == self._STATUS_EXIST:
                    item.setForeground(QColor("#89b4fa"))
                else:
                    item.setForeground(QColor("#f38ba8"))
                self._tbl.setItem(row, col, item)

        self._lbl_summary.setText(
            f"✅ {n_new} nuovi   ℹ️ {n_ex} già presenti   ⚠️ {n_err} errori"
        )
        self._btn_import.setEnabled(n_new > 0)
        self._btn_import.setText(f"Importa validi ({n_new})")

    # ==================================================================
    # Import
    # ==================================================================
    def _do_import(self):
        to_insert = [
            {"code": r["code"], "description": r["desc"]}
            for r in self._preview_rows
            if r["status"] == self._STATUS_NEW
        ]
        if not to_insert:
            return

        try:
            if self._mode == "machines":
                result = session.coding.bulk_import_machines(to_insert)
            else:
                result = session.coding.bulk_import_groups(self._machine_id, to_insert)
        except Exception as e:
            QMessageBox.critical(self, "Errore importazione", str(e))
            return

        msg = (
            f"Importazione completata:\n\n"
            f"  Inseriti:  {result['inserted']}\n"
            f"  Saltati:   {result['skipped']}\n"
            f"  Errori:    {len(result['errors'])}"
        )
        if result["errors"]:
            msg += "\n\nErrori:\n" + "\n".join(result["errors"][:10])

        QMessageBox.information(self, "Importazione completata", msg)

        # Rianalizza per aggiornare la preview (i nuovi diventano "esistenti")
        self._run_analysis()

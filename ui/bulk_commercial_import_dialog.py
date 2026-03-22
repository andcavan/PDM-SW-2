# =============================================================================
#  ui/bulk_commercial_import_dialog.py  –  Importazione massiva Categorie/Sottocategorie
# =============================================================================
from __future__ import annotations
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QGroupBox,
    QTableWidget, QTableWidgetItem, QHeaderView, QTextEdit,
    QMessageBox
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor

from ui.session import session


class BulkCommercialImportDialog(QDialog):
    """
    Importazione massiva di Categorie o Sottocategorie commerciali/normalizzate.

    mode = "categories"    → crea categorie per item_type
    mode = "subcategories" → crea sottocategorie per category_id
    """

    _STATUS_NEW   = "✅ Nuovo"
    _STATUS_EXIST = "ℹ️ Esistente"
    _STATUS_ERR   = "⚠️ Errore"

    def __init__(self, mode: str = "categories",
                 item_type: str = "commerciale",
                 category_id: int | None = None,
                 category_code: str = "",
                 parent=None):
        super().__init__(parent)
        self._mode        = mode
        self._item_type   = item_type
        self._category_id = category_id
        self._preview_rows: list[dict] = []

        if mode == "categories":
            tipo = "Commerciali (5)" if item_type == "commerciale" else "Normalizzati (6)"
            title = f"Importa Categorie — {tipo}"
        else:
            title = f"Importa Sottocategorie — Categoria: {category_code}"

        self.setWindowTitle(title)
        self.setMinimumSize(680, 500)
        self._build_ui()

    # ==================================================================
    # UI
    # ==================================================================
    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setSpacing(10)
        root.setContentsMargins(16, 16, 16, 16)

        # Lista libera
        input_grp = QGroupBox("Lista libera")
        input_lyt = QVBoxLayout(input_grp)
        input_lyt.setContentsMargins(12, 12, 12, 12)

        lbl = QLabel(
            "Inserisci una riga per codice nel formato  <b>CODICE ; Descrizione</b>\n"
            "(la descrizione è opzionale — codice max 10 caratteri)"
        )
        lbl.setWordWrap(True)
        input_lyt.addWidget(lbl)

        self._txt_lista = QTextEdit()
        if self._mode == "categories":
            self._txt_lista.setPlaceholderText(
                "VIT ; Viteria\n"
                "BUL ; Bulloneria\n"
                "RON"
            )
        else:
            self._txt_lista.setPlaceholderText(
                "ISO ; Metriche ISO\n"
                "UNI ; Norme UNI\n"
                "DIN"
            )
        input_lyt.addWidget(self._txt_lista)
        root.addWidget(input_grp)

        # Pulsante analizza
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

    # ==================================================================
    # Logica analisi
    # ==================================================================
    def _run_analysis(self):
        rows = self._rows_from_lista()
        if not rows:
            QMessageBox.warning(self, "Nessun dato", "Nessun codice da analizzare.")
            return

        # Codici già esistenti
        if self._mode == "categories":
            existing = {
                c["code"]
                for c in session.commercial.get_categories(item_type=self._item_type)
            }
        else:
            existing = {
                s["code"]
                for s in session.commercial.get_subcategories(self._category_id)
            } if self._category_id else set()

        preview = []
        for code, desc in rows:
            if not code:
                preview.append({"code": code, "desc": desc,
                                 "status": self._STATUS_ERR, "note": "codice vuoto"})
            elif len(code) > 10:
                preview.append({"code": code, "desc": desc,
                                 "status": self._STATUS_ERR,
                                 "note": "codice supera 10 caratteri"})
            elif code in existing:
                preview.append({"code": code, "desc": desc,
                                 "status": self._STATUS_EXIST, "note": "già presente"})
            else:
                preview.append({"code": code, "desc": desc,
                                 "status": self._STATUS_NEW, "note": ""})

        self._preview_rows = preview
        self._refresh_preview()

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

        inserted = 0
        errors = []
        for item in to_insert:
            try:
                if self._mode == "categories":
                    session.commercial.create_category(
                        item["code"], item["description"], self._item_type
                    )
                else:
                    session.commercial.create_subcategory(
                        self._category_id, item["code"], item["description"], ""
                    )
                inserted += 1
            except Exception as e:
                errors.append(f"{item['code']}: {e}")

        msg = (
            f"Importazione completata:\n\n"
            f"  Inseriti:  {inserted}\n"
            f"  Errori:    {len(errors)}"
        )
        if errors:
            msg += "\n\nErrori:\n" + "\n".join(errors[:10])

        QMessageBox.information(self, "Importazione completata", msg)

        # Rianalizza per aggiornare la preview
        self._run_analysis()

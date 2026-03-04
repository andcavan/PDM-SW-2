# =============================================================================
#  ui/checkout_dialog.py  –  Dialog per checkout con opzioni DRW e ASM
# =============================================================================
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QCheckBox, QGroupBox, QMessageBox,
    QTableWidget, QTableWidgetItem, QHeaderView
)
from PyQt6.QtCore import Qt

from ui.session import session
from ui.styles import TYPE_ICON


class CheckoutDialog(QDialog):
    """
    Dialog per confermare il checkout di un documento.
    - Mostra info documento
    - Checkbox "Includi DRW in workspace"
    - Per ASM: info su copia ricorsiva componenti
    """

    def __init__(self, document_id: int, parent=None):
        super().__init__(parent)
        self.document_id = document_id
        self.doc = session.files.get_document(document_id)
        self.result_data = None
        self.setWindowTitle("Checkout")
        self.setMinimumWidth(420)
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        if not self.doc:
            layout.addWidget(QLabel("Documento non trovato"))
            return

        # Info documento
        grp = QGroupBox("Documento")
        grp_layout = QVBoxLayout(grp)

        icon = TYPE_ICON.get(self.doc["doc_type"], "")
        grp_layout.addWidget(QLabel(
            f"{icon}  {self.doc['code']}  Rev. {self.doc['revision']}"
        ))
        grp_layout.addWidget(QLabel(f"Tipo: {self.doc['doc_type']}"))
        grp_layout.addWidget(QLabel(f"Titolo: {self.doc['title'] or ''}"))
        grp_layout.addWidget(QLabel(f"Stato: {self.doc['state']}"))
        layout.addWidget(grp)

        # Opzioni
        grp_opt = QGroupBox("Opzioni checkout")
        opt_layout = QVBoxLayout(grp_opt)

        self.chk_drw = QCheckBox("Includi anche il Disegno (DRW) in workspace")
        self.chk_drw.setChecked(False)

        # Verifica se esiste un DRW associato
        if self.doc["doc_type"] in ("Parte", "Assieme"):
            drw = session.files.get_drw_document(self.document_id)
            if drw:
                from core.checkout_manager import READONLY_STATES
                if drw["state"] in READONLY_STATES:
                    self.chk_drw.setText(
                        f"DRW associato ({drw['state']}) \u2014 non disponibile per checkout"
                    )
                    self.chk_drw.setEnabled(False)
                elif drw["is_locked"]:
                    locker = drw.get("locked_by_name", "altro utente")
                    self.chk_drw.setText(
                        f"DRW gi\u00e0 in checkout da {locker}"
                    )
                    self.chk_drw.setEnabled(False)
            else:
                self.chk_drw.setText("Nessun DRW associato")
                self.chk_drw.setEnabled(False)
            opt_layout.addWidget(self.chk_drw)
        elif self.doc["doc_type"] == "Disegno":
            # DRW diretto, nessuna opzione DRW
            pass

        # Info ASM: tabella componenti
        if self.doc["doc_type"] == "Assieme":
            components = session.asm.get_bom_flat(self.document_id)
            if components:
                grp_comp = QGroupBox(
                    f"Componenti assieme ({len(components)} file)"
                )
                comp_layout = QVBoxLayout(grp_comp)

                lbl = QLabel(
                    "ℹ️  Questi file verranno copiati in workspace "
                    "(senza lock — solo per apertura assieme)"
                )
                lbl.setWordWrap(True)
                lbl.setStyleSheet("color: #89b4fa; font-style: italic;")
                comp_layout.addWidget(lbl)

                tbl = QTableWidget(0, 4)
                tbl.setHorizontalHeaderLabels(["Codice", "Tipo", "Rev.", "Titolo"])
                tbl.horizontalHeader().setSectionResizeMode(
                    3, QHeaderView.ResizeMode.Stretch
                )
                tbl.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
                tbl.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
                tbl.verticalHeader().setVisible(False)
                tbl.setMaximumHeight(200)

                for c in components:
                    r = tbl.rowCount()
                    tbl.insertRow(r)
                    tbl.setItem(r, 0, QTableWidgetItem(c["code"]))
                    icon_c = TYPE_ICON.get(c["doc_type"], "")
                    tbl.setItem(r, 1, QTableWidgetItem(f"{icon_c} {c['doc_type']}"))
                    tbl.setItem(r, 2, QTableWidgetItem(c["revision"]))
                    tbl.setItem(r, 3, QTableWidgetItem(c.get("title", "")))

                comp_layout.addWidget(tbl)
                layout.addWidget(grp_comp)

        layout.addWidget(grp_opt)

        # Bottoni
        btn_row = QHBoxLayout()
        btn_row.addStretch()

        btn_cancel = QPushButton("Annulla")
        btn_cancel.clicked.connect(self.reject)

        btn_ok = QPushButton("\U0001f4e4  Checkout")
        btn_ok.setObjectName("btn_primary")
        btn_ok.clicked.connect(self._do_checkout)

        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_ok)
        layout.addLayout(btn_row)

    def _do_checkout(self):
        try:
            include_drw = self.chk_drw.isChecked() if self.chk_drw.isEnabled() else False

            if self.doc["doc_type"] == "Assieme":
                result = session.checkout.checkout_asm(
                    self.document_id, include_drw=include_drw
                )
                self.result_data = result
                paths = [str(result["asm_path"])]
                paths.extend(str(p) for p in result["component_paths"])
                msg = f"Checkout ASM eseguito.\n\nFile copiati in workspace:\n"
                msg += f"  \u2022 ASM: {result['asm_path'].name}\n"
                msg += f"  \u2022 {len(result['component_paths'])} componenti\n"
            else:
                dest = session.checkout.checkout(
                    self.document_id, include_drw=include_drw
                )
                self.result_data = {"path": dest}
                msg = f"Checkout eseguito.\n\nFile: {dest.name}"

            QMessageBox.information(self, "Checkout eseguito", msg)
            self.accept()

        except Exception as e:
            QMessageBox.critical(self, "Errore checkout", str(e))

# =============================================================================
#  ui/document_dialog.py  –  Creazione / modifica documento
# =============================================================================
from pathlib import Path
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QComboBox, QTextEdit, QGroupBox, QFormLayout,
    QTabWidget, QWidget, QTableWidget, QTableWidgetItem,
    QMessageBox, QFileDialog, QHeaderView
)
from PyQt6.QtCore import Qt

from typing import Optional

from config import SW_EXTENSIONS
from ui.session import session
from ui.styles import STATE_BADGE_STYLE, TYPE_ICON


class DocumentDialog(QDialog):
    """Dialogo per creare o visualizzare/modificare un documento."""

    def __init__(self, document_id: Optional[int] = None, parent=None):
        super().__init__(parent)
        self.document_id = document_id
        self.is_new      = document_id is None
        self.setWindowTitle("Nuovo Documento" if self.is_new else "Dettaglio Documento")
        self.setMinimumSize(680, 560)
        self._build_ui()
        if not self.is_new:
            self._load_document()

    # ------------------------------------------------------------------
    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(8)
        layout.setContentsMargins(16, 16, 16, 16)

        # Tab
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)

        self.tabs.addTab(self._build_general_tab(), "Generale")
        self.tabs.addTab(self._build_properties_tab(), "Proprietà SW")
        if not self.is_new:
            self.tabs.addTab(self._build_bom_tab(), "Struttura (BOM)")
            self.tabs.addTab(self._build_history_tab(), "Storico")

        # Bottoni
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        if self.is_new:
            btn_save = QPushButton("Crea Documento")
            btn_save.setObjectName("btn_primary")
            btn_save.clicked.connect(self._create)
        else:
            btn_save = QPushButton("Salva Modifiche")
            btn_save.setObjectName("btn_primary")
            btn_save.clicked.connect(self._save)
        btn_close = QPushButton("Chiudi")
        btn_close.clicked.connect(self.accept)
        btn_row.addWidget(btn_close)
        btn_row.addWidget(btn_save)
        layout.addLayout(btn_row)

    # ---- Tab GENERALE ----
    def _build_general_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setSpacing(10)

        grp = QGroupBox("Informazioni documento")
        form = QFormLayout(grp)

        if self.is_new:
            # ---- Selezione gerarchica (solo creazione) ----
            self.cmb_machine = QComboBox()
            self.cmb_machine.setMinimumWidth(180)
            self.cmb_machine.currentIndexChanged.connect(self._on_machine_changed)
            form.addRow("Macchina:", self.cmb_machine)

            self.cmb_level = QComboBox()
            self.cmb_level.addItem("LIV0 – Macchina (ASM)",   0)
            self.cmb_level.addItem("LIV1 – Gruppo   (ASM)",   1)
            self.cmb_level.addItem("LIV2 – Parte    (PRT)",   2)
            self.cmb_level.addItem("LIV2 – Sottogruppo (ASM)", 3)
            self.cmb_level.addItem("Disegno  (DRW)",            4)
            self.cmb_level.currentIndexChanged.connect(self._on_level_changed)
            form.addRow("Livello:", self.cmb_level)

            self.cmb_group = QComboBox()
            self.cmb_group.setMinimumWidth(180)
            self.cmb_group.currentIndexChanged.connect(self._update_code_preview)
            self._lbl_group = form.labelForField(self.cmb_group) if False else QLabel("Gruppo:")
            form.addRow("Gruppo:", self.cmb_group)

            # Documento padre (solo per DRW)
            self.cmb_parent = QComboBox()
            self.cmb_parent.setMinimumWidth(220)
            self.cmb_parent.setVisible(False)
            form.addRow("Doc. padre (PRT/ASM):", self.cmb_parent)

            # Anteprima codice auto
            self.lbl_code_preview = QLabel("—")
            self.lbl_code_preview.setStyleSheet(
                "color:#89b4fa;font-weight:bold;font-size:14px;"
            )
            form.addRow("Codice generato:", self.lbl_code_preview)

            # Tipo nascosto (usato internamente)
            self.cmb_type = QComboBox()  # stub per compat.
        else:
            # Visualizzazione documento esistente
            self.lbl_code_preview = QLabel("")
            self.lbl_code_preview.setStyleSheet("font-weight:bold;color:#89b4fa;")
            form.addRow("Codice:", self.lbl_code_preview)
            self.cmb_type = QComboBox()
            self.cmb_type.addItems(["Parte", "Assieme", "Disegno"])

        # Revisione
        self.txt_rev = QLineEdit("00")
        self.txt_rev.setMaximumWidth(60)
        self.txt_rev.setPlaceholderText("00")
        if not self.is_new:
            form.addRow("Tipo:", self.cmb_type)
        form.addRow("Revisione:", self.txt_rev)

        # Titolo
        self.txt_title = QLineEdit()
        self.txt_title.setPlaceholderText("Descrizione breve del documento")
        form.addRow("Titolo:", self.txt_title)

        # Descrizione
        self.txt_desc = QTextEdit()
        self.txt_desc.setMaximumHeight(80)
        self.txt_desc.setPlaceholderText("Note aggiuntive...")
        form.addRow("Descrizione:", self.txt_desc)

        layout.addWidget(grp)

        # Carica macchine se nuovo documento
        if self.is_new:
            self._load_machines_combo()

        # File
        if not self.is_new:
            grp_file = QGroupBox("File")
            file_form = QFormLayout(grp_file)
            self.lbl_file = QLabel("—")
            file_form.addRow("File in archivio:", self.lbl_file)

            file_row = QHBoxLayout()
            btn_create_sw = QPushButton("\u2728 Crea in SW")
            btn_create_sw.setObjectName("btn_primary")
            btn_create_sw.setToolTip(
                "Crea file da template nella workspace e apre in SolidWorks"
            )
            btn_create_sw.clicked.connect(self._create_from_template)

            btn_copy_file = QPushButton("\U0001f4cb Crea da file")
            btn_copy_file.setToolTip(
                "Copia un file esistente nella workspace rinominandolo col codice PDM"
            )
            btn_copy_file.clicked.connect(self._copy_from_file)

            btn_import = QPushButton("\U0001f4e5 Importa in PDM")
            btn_import.setToolTip(
                "Cerca il file nella workspace e lo archivia nel PDM"
            )
            btn_import.clicked.connect(self._import_from_pdm)

            btn_export = QPushButton("\U0001f4e4 Esporta in WS")
            btn_export.setToolTip(
                "Copia il file dall'archivio nella workspace locale"
            )
            btn_export.clicked.connect(self._export_file)

            btn_open = QPushButton("\U0001f527 Apri in SW")
            btn_open.setToolTip(
                "Copia dall'archivio in workspace (se necessario) e apre in SolidWorks"
            )
            btn_open.clicked.connect(self._open_sw)

            file_row.addWidget(btn_create_sw)
            file_row.addWidget(btn_copy_file)
            file_row.addWidget(btn_import)
            file_row.addWidget(btn_export)
            file_row.addWidget(btn_open)
            file_row.addStretch()
            file_form.addRow(file_row)
            layout.addWidget(grp_file)

            # Stato
            grp_state = QGroupBox("Stato workflow")
            state_form = QFormLayout(grp_state)
            self.lbl_state = QLabel("—")
            self.lbl_locked = QLabel("—")
            state_form.addRow("Stato:", self.lbl_state)
            state_form.addRow("Checkout:", self.lbl_locked)
            layout.addWidget(grp_state)

        layout.addStretch()
        return w

    # ---- Tab PROPRIETÀ ----
    def _build_properties_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)

        btn_row = QHBoxLayout()
        btn_add  = QPushButton("+ Aggiungi")
        btn_add.clicked.connect(self._add_property_row)
        btn_del  = QPushButton("- Rimuovi")
        btn_del.clicked.connect(self._remove_property_row)
        btn_import_sw = QPushButton("Importa da SW")
        btn_import_sw.clicked.connect(self._import_props_from_sw)
        btn_export_xl = QPushButton("Esporta Excel")
        btn_export_xl.clicked.connect(self._export_props_excel)
        btn_import_xl = QPushButton("Importa Excel")
        btn_import_xl.clicked.connect(self._import_props_excel)
        btn_save_props = QPushButton("Salva Proprietà")
        btn_save_props.setObjectName("btn_primary")
        btn_save_props.clicked.connect(self._save_properties)
        for b in [btn_add, btn_del, btn_import_sw,
                  btn_export_xl, btn_import_xl, btn_save_props]:
            btn_row.addWidget(b)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        self.tbl_props = QTableWidget(0, 2)
        self.tbl_props.setHorizontalHeaderLabels(["Proprietà", "Valore"])
        self.tbl_props.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeMode.Stretch
        )
        layout.addWidget(self.tbl_props)
        return w

    # ---- Tab BOM ----
    def _build_bom_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)

        btn_row = QHBoxLayout()
        btn_add_comp  = QPushButton("+ Aggiungi componente")
        btn_add_comp.clicked.connect(self._add_component)
        btn_del_comp  = QPushButton("- Rimuovi")
        btn_del_comp.clicked.connect(self._del_component)
        btn_import_asm = QPushButton("Importa struttura da SW")
        btn_import_asm.clicked.connect(self._import_asm_from_sw)
        for b in [btn_add_comp, btn_del_comp, btn_import_asm]:
            btn_row.addWidget(b)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        self.tbl_bom = QTableWidget(0, 5)
        self.tbl_bom.setHorizontalHeaderLabels(
            ["Codice", "Rev.", "Tipo", "Titolo", "Qtà"]
        )
        self.tbl_bom.horizontalHeader().setSectionResizeMode(
            3, QHeaderView.ResizeMode.Stretch
        )
        layout.addWidget(self.tbl_bom)
        return w

    # ---- Tab STORICO ----
    def _build_history_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        self.tbl_history = QTableWidget(0, 5)
        self.tbl_history.setHorizontalHeaderLabels(
            ["Data/Ora", "Azione", "Da", "A", "Utente"]
        )
        self.tbl_history.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch
        )
        layout.addWidget(self.tbl_history)
        return w

    # ------------------------------------------------------------------
    # Logica
    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # Gestori selezione gerarchica (solo creazione)
    # ------------------------------------------------------------------
    def _load_machines_combo(self):
        machines = session.coding.get_machines(only_active=True)
        self.cmb_machine.blockSignals(True)
        self.cmb_machine.clear()
        for m in machines:
            self.cmb_machine.addItem(
                f"{m['code']} – {m['description'] or ''}", m["id"]
            )
        self.cmb_machine.blockSignals(False)
        self._on_machine_changed()

    def _on_machine_changed(self):
        machine_id = self.cmb_machine.currentData()
        # Aggiorna combo gruppo
        groups = session.coding.get_groups(machine_id) if machine_id else []
        self.cmb_group.blockSignals(True)
        self.cmb_group.clear()
        for g in groups:
            self.cmb_group.addItem(
                f"{g['code']} – {g['description'] or ''}", g["id"]
            )
        self.cmb_group.blockSignals(False)
        self._update_code_preview()

    def _on_level_changed(self):
        level = self.cmb_level.currentData()
        # Nascondi gruppo solo per LIV0 (macchina)
        self.cmb_group.setVisible(level != 0)
        # Mostra parent solo per DRW
        self.cmb_parent.setVisible(level == 4)
        if level == 4:
            self._load_parent_combo()
        self._update_code_preview()

    def _load_parent_combo(self):
        """Popola il combo documento padre con PRT e ASM disponibili."""
        self.cmb_parent.clear()
        self.cmb_parent.addItem("(nessuno)", None)
        docs = session.files.search_documents()
        for d in docs:
            if d["doc_type"] in ("Parte", "Assieme"):
                self.cmb_parent.addItem(
                    f"{d['code']}  rev.{d['revision']}  –  {d['title']}",
                    d["id"]
                )

    def _update_code_preview(self):
        if not self.is_new:
            return
        machine_id = self.cmb_machine.currentData()
        group_id   = self.cmb_group.currentData()
        level_data = self.cmb_level.currentData()
        level_map  = {0: (0, "ASM"), 1: (1, "ASM"), 2: (2, "PRT"), 3: (2, "ASM"), 4: (2, "PRT")}
        level, subtype = level_map.get(level_data, (2, "PRT"))
        if not machine_id:
            self.lbl_code_preview.setText("—")
            return
        try:
            preview = session.coding.preview_code(
                level, subtype, machine_id,
                group_id if level > 0 else None
            )
            self.lbl_code_preview.setText(preview)
        except Exception:
            self.lbl_code_preview.setText("—")

    def _gen_code(self):
        """Stub mantenuto per compatibilità."""
        pass

    # ------------------------------------------------------------------
    def _create(self):
        revision   = self.txt_rev.text().strip() or "00"
        title      = self.txt_title.text().strip()
        desc       = self.txt_desc.toPlainText().strip()
        machine_id = self.cmb_machine.currentData()
        group_id   = self.cmb_group.currentData()
        level_data = self.cmb_level.currentData()

        if not title:
            QMessageBox.warning(self, "Errore", "Il titolo è obbligatorio")
            return
        if not machine_id:
            QMessageBox.warning(self, "Errore", "Seleziona una macchina")
            return
        if level_data in (1, 2, 3, 4) and not group_id:
            QMessageBox.warning(self, "Errore", "Seleziona un gruppo")
            return

        # Per DRW: parent è facoltativo
        parent_doc_id = None
        if level_data == 4:
            parent_doc_id = self.cmb_parent.currentData()  # può essere None
            group_id = self.cmb_group.currentData() if self.cmb_group.isVisible() else None

        # Mappa livello → (func, doc_type, doc_level)
        try:
            if level_data == 0:
                code      = session.coding.next_code_liv0(machine_id)
                doc_type  = "Assieme"
                doc_level = 0
                group_id  = None
            elif level_data == 1:
                code      = session.coding.next_code_liv1(machine_id, group_id)
                doc_type  = "Assieme"
                doc_level = 1
            elif level_data == 2:
                code      = session.coding.next_code_liv2_part(machine_id, group_id)
                doc_type  = "Parte"
                doc_level = 2
            elif level_data == 3:
                code      = session.coding.next_code_liv2_subgroup(machine_id, group_id)
                doc_type  = "Assieme"
                doc_level = 2
            else:  # 4 = Disegno DRW
                code      = session.coding.next_code_liv2_part(machine_id, group_id)
                doc_type  = "Disegno"
                doc_level = 2
        except Exception as e:
            QMessageBox.critical(self, "Errore generazione codice", str(e))
            return

        if not session.coding.is_code_available(code, revision, doc_type):
            QMessageBox.warning(
                self, "Codice duplicato",
                f"Il codice {code} rev.{revision} [{doc_type}] è già presente"
            )
            return

        doc_id = session.files.create_document(
            code, revision, doc_type, title, desc,
            machine_id=machine_id, group_id=group_id,
            doc_level=doc_level, parent_doc_id=parent_doc_id
        )
        self.document_id = doc_id
        self.is_new = False

        QMessageBox.information(
            self, "Documento Creato",
            f"Documento creato:\n{code}  rev.{revision}  [{doc_type}]"
        )
        self.accept()

    def _save(self):
        title = self.txt_title.text().strip()
        desc  = self.txt_desc.toPlainText().strip()
        if not title:
            QMessageBox.warning(self, "Errore", "Il titolo è obbligatorio")
            return
        session.files.update_document(self.document_id, title, desc)
        QMessageBox.information(self, "OK", "Documento aggiornato")

    # ------------------------------------------------------------------
    def _load_document(self):
        doc = session.files.get_document(self.document_id)
        if not doc:
            return
        idx = self.cmb_type.findText(doc["doc_type"])
        if idx >= 0:
            self.cmb_type.setCurrentIndex(idx)
        self.lbl_code_preview.setText(doc["code"])
        self.txt_rev.setText(doc["revision"])
        self.txt_rev.setReadOnly(True)
        self.txt_title.setText(doc["title"])
        self.txt_desc.setPlainText(doc["description"] or "")

        state    = doc["state"]
        badge    = STATE_BADGE_STYLE.get(state, "")
        self.lbl_state.setText(f"<span style='{badge}'>{state}</span>")
        self.lbl_state.setTextFormat(Qt.TextFormat.RichText)

        if doc["is_locked"]:
            locked_name = doc.get("locked_by_name", "sconosciuto")
            self.lbl_locked.setText(
                f"⚠️  In checkout: {locked_name} ({doc.get('locked_ws','')})"
            )
            self.lbl_locked.setStyleSheet("color:#fab387;")
        else:
            self.lbl_locked.setText("✅  Disponibile")
            self.lbl_locked.setStyleSheet("color:#a6e3a1;")

        if doc.get("file_name"):
            self.lbl_file.setText(
                doc["file_name"] + (f"  [{doc['archive_path']}]"
                                    if doc.get("archive_path") else "  (non archiviato)")
            )

        # Proprietà
        props = session.properties.get_properties(self.document_id)
        self.tbl_props.setRowCount(0)
        for name, value in props.items():
            row = self.tbl_props.rowCount()
            self.tbl_props.insertRow(row)
            self.tbl_props.setItem(row, 0, QTableWidgetItem(name))
            self.tbl_props.setItem(row, 1, QTableWidgetItem(value))

        # BOM
        self._refresh_bom()
        # Storico
        self._refresh_history()

    # ------------------------------------------------------------------
    def _refresh_bom(self):
        if not hasattr(self, "tbl_bom"):
            return
        comps = session.asm.get_components(self.document_id)
        self.tbl_bom.setRowCount(0)
        for c in comps:
            row = self.tbl_bom.rowCount()
            self.tbl_bom.insertRow(row)
            self.tbl_bom.setItem(row, 0, QTableWidgetItem(c["code"]))
            self.tbl_bom.setItem(row, 1, QTableWidgetItem(c["revision"]))
            self.tbl_bom.setItem(
                row, 2, QTableWidgetItem(
                    TYPE_ICON.get(c["doc_type"], "") + " " + c["doc_type"]
                )
            )
            self.tbl_bom.setItem(row, 3, QTableWidgetItem(c["title"]))
            self.tbl_bom.setItem(row, 4, QTableWidgetItem(str(c["quantity"])))
            self.tbl_bom.item(row, 4).setData(Qt.ItemDataRole.UserRole, c["child_id"])

    def _refresh_history(self):
        if not hasattr(self, "tbl_history"):
            return
        rows = session.workflow.get_history(self.document_id)
        rows += session.checkout.get_log(self.document_id) if session.checkout else []
        rows.sort(key=lambda r: r.get("changed_at") or r.get("timestamp", ""), reverse=True)
        self.tbl_history.setRowCount(0)
        for r in rows:
            i = self.tbl_history.rowCount()
            self.tbl_history.insertRow(i)
            ts     = r.get("changed_at") or r.get("timestamp", "")
            action = r.get("action", "cambio stato")
            frm    = r.get("from_state", "")
            to     = r.get("to_state", "")
            user   = r.get("user_name", "")
            self.tbl_history.setItem(i, 0, QTableWidgetItem(ts[:19]))
            self.tbl_history.setItem(i, 1, QTableWidgetItem(action))
            self.tbl_history.setItem(i, 2, QTableWidgetItem(frm))
            self.tbl_history.setItem(i, 3, QTableWidgetItem(to))
            self.tbl_history.setItem(i, 4, QTableWidgetItem(user))

    # ------------------------------------------------------------------
    # Proprietà
    def _add_property_row(self):
        row = self.tbl_props.rowCount()
        self.tbl_props.insertRow(row)
        self.tbl_props.setItem(row, 0, QTableWidgetItem(""))
        self.tbl_props.setItem(row, 1, QTableWidgetItem(""))
        self.tbl_props.editItem(self.tbl_props.item(row, 0))

    def _remove_property_row(self):
        row = self.tbl_props.currentRow()
        if row >= 0:
            self.tbl_props.removeRow(row)

    def _save_properties(self):
        if not self.document_id:
            return
        props = {}
        for i in range(self.tbl_props.rowCount()):
            name = self.tbl_props.item(i, 0)
            val  = self.tbl_props.item(i, 1)
            if name and name.text().strip():
                props[name.text().strip()] = val.text() if val else ""
        session.properties.save_properties(self.document_id, props)
        QMessageBox.information(self, "OK", f"{len(props)} proprietà salvate")

    def _import_props_from_sw(self):
        if not self.document_id:
            return
        doc = session.files.get_document(self.document_id)
        if not doc or not doc.get("file_name"):
            QMessageBox.warning(self, "Errore", "Nessun file associato al documento")
            return

        # Costruisci path archivio (usato come fallback da read_from_sw_file)
        archive_path = None
        if doc.get("archive_path"):
            archive_path = session.checkout.sp.root / doc["archive_path"]

        try:
            # Legge proprietà da SolidWorks cercando prima il doc aperto per nome
            props = session.properties.read_from_sw_file(
                archive_path or Path(doc["file_name"]),
                file_name=doc["file_name"],
            )

            # Controlla errore COM
            err = props.pop("_error", None)
            if err:
                QMessageBox.warning(
                    self, "Errore lettura SolidWorks",
                    f"Impossibile leggere le proprietà da SolidWorks:\n\n{err}\n\n"
                    "Assicurarsi che il file sia aperto in SolidWorks.",
                )
                return

            if not props:
                QMessageBox.information(
                    self, "Nessuna proprietà",
                    "Nessuna proprietà custom trovata nel file SolidWorks.",
                )
                return

            # Salva direttamente nel DB
            session.properties.save_properties(self.document_id, props)

            # Aggiorna tabella: pulisci e ricarica dal DB
            self.tbl_props.setRowCount(0)
            all_props = session.properties.get_properties(self.document_id)
            for name, value in all_props.items():
                row = self.tbl_props.rowCount()
                self.tbl_props.insertRow(row)
                self.tbl_props.setItem(row, 0, QTableWidgetItem(name))
                self.tbl_props.setItem(row, 1, QTableWidgetItem(value))

            QMessageBox.information(
                self, "Importazione completata",
                f"{len(props)} proprietà importate da SolidWorks.",
            )
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

    def _export_props_excel(self):
        if not self.document_id:
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Esporta proprietà", "", "Excel (*.xlsx)"
        )
        if path:
            try:
                session.properties.export_to_excel(
                    self.document_id, Path(path)
                )
                QMessageBox.information(self, "OK", f"Esportato: {path}")
            except Exception as e:
                QMessageBox.critical(self, "Errore", str(e))

    def _import_props_excel(self):
        if not self.document_id:
            return
        path, _ = QFileDialog.getOpenFileName(
            self, "Importa proprietà", "", "Excel (*.xlsx)"
        )
        if path:
            try:
                n = session.properties.import_from_excel(
                    self.document_id, Path(path)
                )
                self._load_document()
                QMessageBox.information(self, "OK", f"{n} proprietà importate")
            except Exception as e:
                QMessageBox.critical(self, "Errore", str(e))

    # ------------------------------------------------------------------
    # Gestione file
    def _get_workspace_or_warn(self) -> "Path | None":
        """Ritorna workspace configurata; se assente mostra warning e ritorna None."""
        from ui.sw_config_dialog import SWConfigDialog
        ws = SWConfigDialog.get_workspace()
        if not ws:
            QMessageBox.warning(
                self, "Workspace non configurata",
                "Nessuna workspace configurata su questa workstation.\n"
                "Aprire Strumenti \u2192 Configurazione SolidWorks e impostare la cartella workspace."
            )
        return ws

    def _copy_from_file(self):
        """
        Crea da file: seleziona un file da qualsiasi cartella,
        lo copia in workspace rinominandolo col codice PDM.
        Per PRT/ASM cerca un DRW companion nella stessa cartella.
        """
        if not self.document_id:
            return
        ws = self._get_workspace_or_warn()
        if not ws:
            return
        doc = session.files.get_document(self.document_id)
        # Filtro per tipo documento
        type_filters = {
            "Parte":   "SolidWorks Parte (*.SLDPRT *.sldprt)",
            "Assieme": "SolidWorks Assieme (*.SLDASM *.sldasm)",
            "Disegno": "SolidWorks Disegno (*.SLDDRW *.slddrw)",
        }
        exts_filter = type_filters.get(doc["doc_type"] if doc else "", "File SolidWorks (*.SLDPRT *.SLDASM *.SLDDRW)")
        path, _ = QFileDialog.getOpenFileName(
            self, "Seleziona file sorgente", "", exts_filter
        )
        if not path:
            return
        src_path = Path(path)
        try:
            dest = session.files.copy_to_workspace(src_path, self.document_id, ws)
            QMessageBox.information(
                self, "File copiato in workspace",
                f"File copiato in workspace:\n{dest}\n\n"
                "Aprire il file in SolidWorks, eventualmente aggiornare riferimenti,\n"
                "salvare, poi usare 'Importa in PDM' per archiviarlo."
            )
        except ValueError as e:
            QMessageBox.warning(self, "Tipo file non valido", str(e))
            return
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))
            return

        # Per PRT/ASM cerca companion DRW nella stessa cartella
        if doc and doc["doc_type"] in ("Parte", "Assieme"):
            companion = session.files.find_companion_drw(src_path)
            if companion:
                r = QMessageBox.question(
                    self,
                    "Copia DRW associato?",
                    f"Trovato disegno associato:\n{companion.name}\n"
                    "\nCopiarlo in workspace con il codice PDM?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                )
                if r == QMessageBox.StandardButton.Yes:
                    try:
                        drw_doc_id = session.files.get_or_create_drw_document(
                            self.document_id
                        )
                        drw_dest = session.files.copy_to_workspace(
                            companion, drw_doc_id, ws
                        )
                        QMessageBox.information(
                            self, "DRW copiato",
                            f"File DRW copiato in workspace:\n{drw_dest}"
                        )
                    except Exception as e:
                        QMessageBox.critical(self, "Errore DRW", str(e))

    def _import_from_pdm(self):
        """
        Importa in PDM: cerca nella workspace il file con il codice PDM
        e lo archivia. Se non trovato mostra avviso.
        """
        if not self.document_id:
            return
        ws = self._get_workspace_or_warn()
        if not ws:
            return
        doc = session.files.get_document(self.document_id)
        try:
            dest = session.files.import_from_workspace(self.document_id, ws)
            QMessageBox.information(
                self, "Importato in PDM",
                f"File archiviato nel PDM:\n{dest}"
            )
            self._load_document()
        except FileNotFoundError as e:
            QMessageBox.warning(self, "File non trovato in workspace", str(e))
            return
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))
            return

        # Per PRT/ASM: cerca DRW in workspace e proponi importazione
        if doc and doc["doc_type"] in ("Parte", "Assieme"):
            try:
                drw_doc_id = session.files.get_or_create_drw_document(
                    self.document_id
                )
                drw_dest = session.files.import_from_workspace(drw_doc_id, ws)
                QMessageBox.information(
                    self, "DRW importato in PDM",
                    f"File DRW archiviato:\n{drw_dest}"
                )
            except FileNotFoundError:
                pass  # DRW non presente in workspace, non è un errore
            except Exception as e:
                QMessageBox.critical(self, "Errore DRW", str(e))

    def _export_file(self):
        """Esporta dall'archivio alla workspace locale."""
        if not self.document_id:
            return
        ws = self._get_workspace_or_warn()
        if not ws:
            return
        doc = session.files.get_document(self.document_id)
        # Proponi DRW solo se il documento è PRT o ASM
        include_drw = False
        if doc and doc["doc_type"] in ("Parte", "Assieme"):
            drw_doc = session.files.get_drw_document(self.document_id)
            if drw_doc and drw_doc.get("archive_path"):
                r = QMessageBox.question(
                    self,
                    "Esporta DRW?",
                    f"Presente disegno associato ({drw_doc['code']}.SLDDRW).\n"
                    "Esportarlo insieme al file principale nella workspace?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                )
                include_drw = (r == QMessageBox.StandardButton.Yes)
        try:
            exported = session.files.export_to_workspace(
                self.document_id, ws, include_drw=include_drw
            )
            if exported:
                msg = "\n".join(str(p) for p in exported)
                QMessageBox.information(
                    self, "File in workspace",
                    f"File copiati nella workspace:\n{msg}"
                )
            else:
                QMessageBox.warning(
                    self, "Nessun file",
                    "Nessun file in archivio da esportare.\n"
                    "Usare prima 'Importa in PDM'."
                )
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

    def _open_sw(self):
        """Copia dall'archivio in workspace (se necessario) e apre in SolidWorks."""
        if not self.document_id:
            return
        ws = self._get_workspace_or_warn()
        if not ws:
            return
        try:
            session.files.open_from_workspace(self.document_id, ws)
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

    def _create_from_template(self):
        """Crea il file SW da template nella workspace e lo apre."""
        if not self.document_id:
            return
        import os
        ws = self._get_workspace_or_warn()
        if not ws:
            return

        # Per PRT/ASM: chiedi se creare anche il DRW
        doc = session.files.get_document(self.document_id)
        also_drw = False
        if doc and doc["doc_type"] in ("Parte", "Assieme"):
            from ui.sw_config_dialog import SWConfigDialog
            drw_tpl = SWConfigDialog.get_template("Disegno")
            if drw_tpl and drw_tpl.exists():
                r = QMessageBox.question(
                    self,
                    "Crea DRW?",
                    "Creare anche il file disegno (.SLDDRW) nella workspace?\n"
                    "(verrà creato anche il documento Disegno nel PDM)",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                )
                also_drw = (r == QMessageBox.StandardButton.Yes)

        try:
            dest, drw_dest = session.files.create_from_template(
                self.document_id, ws, also_drw=also_drw
            )
            msg = f"File creato:\n{dest}"
            if drw_dest:
                msg += f"\n\nDRW creato:\n{drw_dest}"
            msg += "\n\nAprire il file principale in SolidWorks?"
            r = QMessageBox.question(
                self, "File creato", msg,
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if r == QMessageBox.StandardButton.Yes:
                os.startfile(str(dest))
        except FileNotFoundError as e:
            QMessageBox.warning(self, "Template mancante", str(e))
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

    # ------------------------------------------------------------------
    # BOM
    def _add_component(self):
        from ui.document_selector import DocumentSelectorDialog
        dlg = DocumentSelectorDialog(parent=self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            child_id = dlg.selected_id
            if child_id and child_id != self.document_id:
                session.asm.add_component(self.document_id, child_id, 1.0)
                self._refresh_bom()

    def _del_component(self):
        row = self.tbl_bom.currentRow()
        if row < 0:
            return
        child_id = self.tbl_bom.item(row, 4)
        if child_id:
            cid = child_id.data(Qt.ItemDataRole.UserRole)
            session.asm.remove_component(self.document_id, cid)
            self._refresh_bom()

    def _import_asm_from_sw(self):
        if not self.document_id:
            return
        doc = session.files.get_document(self.document_id)
        if not doc or doc.get("doc_type") != "Assieme":
            QMessageBox.warning(
                self, "Errore",
                "Questa funzione è disponibile solo per documenti di tipo Assieme."
            )
            return
        try:
            n = session.asm.import_bom_from_active_doc(self.document_id)
            self._refresh_bom()
            QMessageBox.information(
                self, "OK",
                f"Importati {n} componenti dall'assieme attivo in SolidWorks.\n\n"
                "Nota: vengono collegati solo i documenti già presenti nel PDM\n"
                "(codice = nome file senza estensione)."
            )
        except Exception as e:
            QMessageBox.critical(self, "Errore import BOM", str(e))

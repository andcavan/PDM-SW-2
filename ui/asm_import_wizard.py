# =============================================================================
#  ui/asm_import_wizard.py  –  Wizard importazione massiva struttura ASM
# =============================================================================
from __future__ import annotations
import logging
import shutil
import socket
from pathlib import Path
from typing import Optional, TYPE_CHECKING

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QComboBox, QGroupBox, QTableWidget, QTableWidgetItem,
    QHeaderView, QMessageBox, QWidget, QCheckBox, QProgressDialog,
    QAbstractItemView,
)
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QFont, QColor

if TYPE_CHECKING:
    from core.database import Database
    from core.coding_manager import CodingManager
    from core.file_manager import FileManager
    from core.asm_manager import AsmManager
    from config import SharedPaths

# ---------------------------------------------------------------------------
#  Costanti
# ---------------------------------------------------------------------------
_LEVEL_OPTIONS = [
    ("LIV0 – Macchina (ASM)",       "ASM_MACH", 0),
    ("LIV1 – Gruppo   (ASM)",       "ASM_GRP",  1),
    ("LIV2 – Sottogruppo (ASM)",    "ASM_SUB",  2),
    ("LIV2 – Parte    (PRT)",       "PRT",      2),
]
_DEFAULT_LEVEL_ASM = 1   # indice in _LEVEL_OPTIONS per ASM (LIV1)
_DEFAULT_LEVEL_PRT = 3   # indice in _LEVEL_OPTIONS per PRT (LIV2-PRT)

STYLE = """
QDialog, QWidget {
    background-color: #1e1e2e;
    color: #cdd6f4;
    font-family: 'Segoe UI', Arial, sans-serif;
    font-size: 12px;
}
QGroupBox {
    border: 1px solid #45475a;
    border-radius: 5px;
    margin-top: 8px;
    padding-top: 6px;
    color: #89b4fa;
    font-weight: bold;
}
QGroupBox::title { subcontrol-origin: margin; left: 8px; padding: 0 4px; }
QPushButton {
    background-color: #313244;
    color: #cdd6f4;
    border: 1px solid #45475a;
    border-radius: 5px;
    padding: 6px 12px;
    min-width: 100px;
}
QPushButton:hover { background-color: #45475a; border-color: #89b4fa; }
QPushButton:pressed { background-color: #585b70; }
QPushButton:disabled { color: #585b70; border-color: #313244; }
QPushButton#btn_primary {
    background-color: #1e3a5f;
    border-color: #89b4fa;
    color: #cdd6f4;
    font-weight: bold;
}
QPushButton#btn_primary:hover { background-color: #2a4a7f; }
QComboBox {
    background-color: #313244;
    border: 1px solid #45475a;
    border-radius: 4px;
    padding: 3px 6px;
    color: #cdd6f4;
}
QComboBox:focus { border-color: #89b4fa; }
QComboBox::drop-down { border: none; width: 18px; }
QComboBox QAbstractItemView {
    background-color: #313244;
    border: 1px solid #45475a;
    selection-background-color: #45475a;
}
QTableWidget {
    background-color: #181825;
    border: 1px solid #45475a;
    gridline-color: #313244;
    color: #cdd6f4;
}
QTableWidget::item { padding: 2px 4px; }
QTableWidget::item:selected { background-color: #313244; color: #89b4fa; }
QHeaderView::section {
    background-color: #1e1e2e;
    color: #89b4fa;
    border: none;
    border-bottom: 1px solid #45475a;
    padding: 4px 6px;
    font-weight: bold;
}
QCheckBox { spacing: 4px; }
QCheckBox::indicator {
    width: 14px; height: 14px;
    border: 1px solid #45475a;
    border-radius: 3px;
    background: #313244;
}
QCheckBox::indicator:checked { background: #89b4fa; border-color: #89b4fa; }
"""

# Colonne della tabella
_COL_NUM    = 0
_COL_INC    = 1
_COL_NAME   = 2
_COL_TYPE   = 3
_COL_PDM    = 4
_COL_MAC    = 5
_COL_GRP    = 6
_COL_LVL    = 7
_COL_CODE   = 8
_NCOLS      = 9


class AsmImportWizard(QDialog):
    """Wizard per l'importazione massiva di un assieme SolidWorks nel PDM."""

    def __init__(self, asm_file: Optional[Path],
                 db: "Database", sp: "SharedPaths", user: dict,
                 coding: "CodingManager", files: "FileManager",
                 asm_mgr: "AsmManager", parent=None):
        super().__init__(parent)
        self.asm_file = asm_file
        self.db       = db
        self.sp       = sp
        self.user     = user
        self.coding   = coding
        self.files    = files
        self.asm_mgr  = asm_mgr

        # Lista nodi letta da SW: [{name, path, type, depth, parent_path, quantity}]
        self._nodes: list[dict] = []
        # Macchine disponibili: [{id, code, description}]
        self._machines: list[dict] = []

        self.setWindowTitle("PDM-SW  –  Importazione Massiva Struttura ASM")
        self.setMinimumSize(1200, 650)
        self.setStyleSheet(STYLE)
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowMaximizeButtonHint)

        self._build_ui()

        # Carica macchine
        self._machines = self.coding.get_machines(only_active=True)
        self._populate_global_combos()

        # Leggi struttura ASM in modo differito (dopo che la finestra è visibile)
        QTimer.singleShot(200, self._load_asm_tree)

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------
    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        # Titolo
        title_label = QLabel()
        title_label.setStyleSheet("font-size:14px; font-weight:bold; color:#89b4fa;")
        fname = self.asm_file.name if self.asm_file else "(da documento attivo SW)"
        title_label.setText(f"Importazione massiva:  {fname}")
        layout.addWidget(title_label)

        # ---- Impostazioni globali ----
        grp_global = QGroupBox("Impostazioni globali  (applica a tutti i componenti selezionati)")
        global_row = QHBoxLayout(grp_global)
        global_row.setSpacing(12)

        global_row.addWidget(QLabel("Macchina:"))
        self.cmb_g_machine = QComboBox()
        self.cmb_g_machine.setMinimumWidth(180)
        self.cmb_g_machine.currentIndexChanged.connect(self._on_global_machine_changed)
        global_row.addWidget(self.cmb_g_machine)

        global_row.addWidget(QLabel("Gruppo:"))
        self.cmb_g_group = QComboBox()
        self.cmb_g_group.setMinimumWidth(160)
        global_row.addWidget(self.cmb_g_group)

        global_row.addWidget(QLabel("Livello:"))
        self.cmb_g_level = QComboBox()
        for label, _, _ in _LEVEL_OPTIONS:
            self.cmb_g_level.addItem(label)
        global_row.addWidget(self.cmb_g_level)

        btn_apply_all = QPushButton("Applica a tutti i selezionati")
        btn_apply_all.clicked.connect(self._apply_global)
        global_row.addWidget(btn_apply_all)
        global_row.addStretch()
        layout.addWidget(grp_global)

        # ---- Tabella componenti ----
        self.table = QTableWidget(0, _NCOLS)
        self.table.setHorizontalHeaderLabels([
            "#", "Incl.", "Componente (file originale)",
            "Tipo", "Già PDM",
            "Macchina", "Gruppo", "Livello", "Codice proposto",
        ])
        hh = self.table.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        hh.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        hh.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        hh.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        hh.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        hh.setSectionResizeMode(5, QHeaderView.ResizeMode.Interactive)
        hh.setSectionResizeMode(6, QHeaderView.ResizeMode.Interactive)
        hh.setSectionResizeMode(7, QHeaderView.ResizeMode.Interactive)
        hh.setSectionResizeMode(8, QHeaderView.ResizeMode.Interactive)
        self.table.setColumnWidth(5, 160)
        self.table.setColumnWidth(6, 140)
        self.table.setColumnWidth(7, 190)
        self.table.setColumnWidth(8, 160)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet(STYLE + "QTableWidget { alternate-background-color: #1a1a2e; }")
        layout.addWidget(self.table)

        self.lbl_status = QLabel("In attesa di lettura struttura ASM da SolidWorks…")
        self.lbl_status.setStyleSheet("color:#a6adc8; font-size:11px;")
        layout.addWidget(self.lbl_status)

        # ---- Pulsanti ----
        btn_row = QHBoxLayout()
        btn_sel_all   = QPushButton("Seleziona tutti")
        btn_sel_all.clicked.connect(self._select_all)
        btn_desel_all = QPushButton("Deseleziona tutti")
        btn_desel_all.clicked.connect(self._deselect_all)
        btn_row.addWidget(btn_sel_all)
        btn_row.addWidget(btn_desel_all)
        btn_row.addStretch()

        btn_cancel = QPushButton("Annulla")
        btn_cancel.clicked.connect(self.reject)

        self.btn_import = QPushButton("Importa e copia in workspace")
        self.btn_import.setObjectName("btn_primary")
        self.btn_import.setEnabled(False)
        self.btn_import.clicked.connect(self._import)
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(self.btn_import)
        layout.addLayout(btn_row)

    # ------------------------------------------------------------------
    # Combo globali
    # ------------------------------------------------------------------
    def _populate_global_combos(self):
        self.cmb_g_machine.blockSignals(True)
        self.cmb_g_machine.clear()
        for m in self._machines:
            self.cmb_g_machine.addItem(f"{m['code']}  –  {m['description']}", m["id"])
        self.cmb_g_machine.blockSignals(False)
        self._on_global_machine_changed()

    def _on_global_machine_changed(self, _=None):
        mid = self.cmb_g_machine.currentData()
        self.cmb_g_group.clear()
        if mid is None:
            return
        groups = self.coding.get_groups(mid)
        for g in groups:
            self.cmb_g_group.addItem(f"{g['code']}  –  {g['description']}", g["id"])

    # ------------------------------------------------------------------
    # Lettura albero ASM
    # ------------------------------------------------------------------
    def _load_asm_tree(self):
        self.lbl_status.setText("Lettura struttura ASM da SolidWorks in corso…")
        try:
            self._nodes = self.asm_mgr.read_asm_tree_from_active()
        except Exception as e:
            self.lbl_status.setText(f"Errore: {e}")
            QMessageBox.critical(
                self, "Errore lettura struttura",
                f"Impossibile leggere la struttura ASM:\n\n{e}"
            )
            return

        if not self._nodes:
            self.lbl_status.setText("Struttura ASM vuota o non leggibile.")
            return

        self._populate_table()
        self.btn_import.setEnabled(True)
        self.lbl_status.setText(
            f"Trovati {len(self._nodes)} componenti (inclusa la radice). "
            "Assegnare la codifica e cliccare 'Importa e copia in workspace'."
        )

    # ------------------------------------------------------------------
    # Popolamento tabella
    # ------------------------------------------------------------------
    def _populate_table(self):
        self.table.setRowCount(0)
        for i, node in enumerate(self._nodes):
            self.table.insertRow(i)
            self._set_row(i, node)
        # Forza aggiornamento codici per tutte le righe dopo la creazione completa
        for i in range(len(self._nodes)):
            self._update_row_code(i)

    def _set_row(self, row: int, node: dict):
        depth = node["depth"]
        indent = "  " * depth + ("└─ " if depth > 0 else "")
        node_type = node["type"]

        # Col 0: numero riga
        num_item = QTableWidgetItem(str(row + 1))
        num_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.table.setItem(row, _COL_NUM, num_item)

        # Col 1: checkbox includi
        # La radice (depth=0) è obbligatoria: checkbox disabilitata e sempre attiva
        chk_widget = QWidget()
        chk_layout = QHBoxLayout(chk_widget)
        chk_layout.setContentsMargins(4, 0, 4, 0)
        chk_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        chk = QCheckBox()
        chk.setChecked(True)
        if depth == 0:
            chk.setEnabled(False)
            chk_widget.setToolTip("L'assieme radice deve sempre essere incluso")
        else:
            chk.stateChanged.connect(lambda state, r=row: self._on_check_changed(r, state))
        chk_layout.addWidget(chk)
        self.table.setCellWidget(row, _COL_INC, chk_widget)

        # Col 2: nome componente
        name_item = QTableWidgetItem(indent + node["name"])
        if depth == 0:
            name_item.setForeground(QColor("#f5c2e7"))
            name_item.setToolTip("Radice ASM – codificato come LIV0 Macchina")
        self.table.setItem(row, _COL_NAME, name_item)

        # Col 3: tipo
        type_icon = "🔩" if node_type == "ASM" else "⚙"
        type_item = QTableWidgetItem(f"{type_icon} {node_type}")
        type_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.table.setItem(row, _COL_TYPE, type_item)

        # Col 4: già in PDM?
        existing = self._check_existing(node)
        node["_existing_id"] = existing["id"] if existing else None
        pdm_text = f"✅ {existing['code']}" if existing else "—"
        pdm_item = QTableWidgetItem(pdm_text)
        pdm_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        if existing:
            pdm_item.setForeground(QColor("#a6e3a1"))
        self.table.setItem(row, _COL_PDM, pdm_item)

        # Col 5: combo macchina
        cmb_mac = QComboBox()
        for m in self._machines:
            cmb_mac.addItem(f"{m['code']}  –  {m['description']}", m["id"])
        # Pre-seleziona valore globale
        g_mid = self.cmb_g_machine.currentData()
        for idx in range(cmb_mac.count()):
            if cmb_mac.itemData(idx) == g_mid:
                cmb_mac.setCurrentIndex(idx)
                break
        self.table.setCellWidget(row, _COL_MAC, cmb_mac)

        # Col 6: combo gruppo
        cmb_grp = QComboBox()
        self._fill_group_combo(cmb_grp, g_mid)
        g_gid = self.cmb_g_group.currentData()
        for idx in range(cmb_grp.count()):
            if cmb_grp.itemData(idx) == g_gid:
                cmb_grp.setCurrentIndex(idx)
                break
        self.table.setCellWidget(row, _COL_GRP, cmb_grp)

        # Col 7: combo livello — default per tipo: radice→LIV0, ASM→LIV1, PRT→LIV2-PRT
        # L'utente può sempre cambiare il livello, anche per la radice
        cmb_lvl = QComboBox()
        for label, _, _ in _LEVEL_OPTIONS:
            cmb_lvl.addItem(label)
        if depth == 0:
            default_lvl = 0                     # LIV0 – Macchina (default radice)
        elif node_type == "ASM":
            default_lvl = _DEFAULT_LEVEL_ASM    # LIV1 – Gruppo
        else:
            default_lvl = _DEFAULT_LEVEL_PRT    # LIV2 – Parte
        cmb_lvl.setCurrentIndex(default_lvl)
        self.table.setCellWidget(row, _COL_LVL, cmb_lvl)

        # Col 8: codice proposto (label aggiornata dai segnali)
        code_item = QTableWidgetItem("—")
        code_item.setForeground(QColor("#89b4fa"))
        code_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.table.setItem(row, _COL_CODE, code_item)

        # Collega segnali per aggiornare anteprima codice
        cmb_mac.currentIndexChanged.connect(
            lambda _, r=row: self._on_row_machine_changed(r)
        )
        cmb_grp.currentIndexChanged.connect(
            lambda _, r=row: self._update_row_code(r)
        )
        cmb_lvl.currentIndexChanged.connect(
            lambda _, r=row: self._update_row_code(r)
        )

        self._update_row_code(row)

    def _fill_group_combo(self, combo: QComboBox, machine_id):
        combo.blockSignals(True)
        combo.clear()
        if machine_id is not None:
            groups = self.coding.get_groups(machine_id)
            for g in groups:
                combo.addItem(f"{g['code']}  –  {g['description']}", g["id"])
            if not groups:
                combo.addItem("(nessun gruppo)", None)
        combo.blockSignals(False)

    # ------------------------------------------------------------------
    # Segnali riga
    # ------------------------------------------------------------------
    def _on_check_changed(self, row: int, state: int):
        checked = (state == Qt.CheckState.Checked.value)
        for col in (_COL_MAC, _COL_GRP, _COL_LVL):
            w = self.table.cellWidget(row, col)
            if w:
                w.setEnabled(checked)

    def _on_row_machine_changed(self, row: int):
        cmb_mac = self.table.cellWidget(row, _COL_MAC)
        cmb_grp = self.table.cellWidget(row, _COL_GRP)
        if cmb_mac and cmb_grp:
            self._fill_group_combo(cmb_grp, cmb_mac.currentData())
        self._update_row_code(row)

    def _update_row_code(self, row: int):
        """Aggiorna l'anteprima del codice nella colonna _COL_CODE per la riga `row`."""
        node = self._nodes[row] if row < len(self._nodes) else None
        if node and node.get("_existing_id"):
            # Già in PDM: mostra il codice esistente
            existing = self.db.fetchone(
                "SELECT code FROM documents WHERE id=?", (node["_existing_id"],)
            )
            code_text = existing["code"] if existing else "già nel PDM"
            item = self.table.item(row, _COL_CODE)
            if item:
                item.setText(code_text)
                item.setForeground(QColor("#a6e3a1"))
            return

        cmb_mac = self.table.cellWidget(row, _COL_MAC)
        cmb_grp = self.table.cellWidget(row, _COL_GRP)
        cmb_lvl = self.table.cellWidget(row, _COL_LVL)
        if not (cmb_mac and cmb_grp and cmb_lvl):
            return

        mid = cmb_mac.currentData()
        gid = cmb_grp.currentData()
        lvl_idx = cmb_lvl.currentIndex()
        if lvl_idx < 0 or lvl_idx >= len(_LEVEL_OPTIONS):
            return
        _, subtype, level = _LEVEL_OPTIONS[lvl_idx]

        item = self.table.item(row, _COL_CODE)
        if not item:
            return

        if not mid:
            item.setText("—")
            return

        try:
            preview = self.coding.preview_code(
                level, "ASM" if "ASM" in subtype else "PRT",
                mid, gid if level > 0 else None,
            )
            item.setText(preview)
            item.setForeground(QColor("#89b4fa"))
        except Exception as e:
            item.setText(f"ERR: {e}")
            item.setForeground(QColor("#f38ba8"))

    # ------------------------------------------------------------------
    # Azioni globali
    # ------------------------------------------------------------------
    def _apply_global(self):
        g_mid    = self.cmb_g_machine.currentData()
        g_gid    = self.cmb_g_group.currentData()
        g_lvl    = self.cmb_g_level.currentIndex()

        for row in range(self.table.rowCount()):
            if not self._is_row_checked(row):
                continue

            cmb_mac = self.table.cellWidget(row, _COL_MAC)
            cmb_grp = self.table.cellWidget(row, _COL_GRP)
            cmb_lvl = self.table.cellWidget(row, _COL_LVL)
            if not (cmb_mac and cmb_grp and cmb_lvl):
                continue

            # Macchina
            for idx in range(cmb_mac.count()):
                if cmb_mac.itemData(idx) == g_mid:
                    cmb_mac.blockSignals(True)
                    cmb_mac.setCurrentIndex(idx)
                    cmb_mac.blockSignals(False)
                    break
            # Ricarica gruppi
            self._fill_group_combo(cmb_grp, g_mid)
            # Gruppo
            for idx in range(cmb_grp.count()):
                if cmb_grp.itemData(idx) == g_gid:
                    cmb_grp.blockSignals(True)
                    cmb_grp.setCurrentIndex(idx)
                    cmb_grp.blockSignals(False)
                    break
            # Livello
            cmb_lvl.blockSignals(True)
            cmb_lvl.setCurrentIndex(g_lvl)
            cmb_lvl.blockSignals(False)
            # Aggiorna anteprima
            self._update_row_code(row)

    def _select_all(self):
        for row in range(self.table.rowCount()):
            self._set_row_checked(row, True)

    def _deselect_all(self):
        for row in range(self.table.rowCount()):
            self._set_row_checked(row, False)

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def _is_row_checked(self, row: int) -> bool:
        w = self.table.cellWidget(row, _COL_INC)
        if w:
            chk = w.findChild(QCheckBox)
            if chk:
                return chk.isChecked()
        return False

    def _set_row_checked(self, row: int, checked: bool):
        w = self.table.cellWidget(row, _COL_INC)
        if w:
            chk = w.findChild(QCheckBox)
            if chk:
                chk.setChecked(checked)

    def _check_existing(self, node: dict) -> Optional[dict]:
        """Cerca nel DB un documento con lo stesso stem del file."""
        stem = Path(node["path"]).stem if node["path"] else node["name"]
        ext  = Path(node["path"]).suffix.upper() if node["path"] else ""
        from config import SW_EXTENSIONS
        doc_type = SW_EXTENSIONS.get(ext, "")
        if not doc_type:
            doc_type_filter = ""
        else:
            doc_type_filter = f"AND doc_type='{doc_type}'"
        return self.db.fetchone(
            f"SELECT * FROM documents WHERE code=? {doc_type_filter} "
            "AND state != 'Obsoleto' ORDER BY revision DESC",
            (stem,),
        )

    # ------------------------------------------------------------------
    # IMPORTAZIONE
    # ------------------------------------------------------------------
    def _import(self):
        """Orchestrate l'intera operazione di importazione."""
        # 1) Raccogli righe selezionate con codice proposto
        rows_to_import: list[dict] = []
        for row in range(self.table.rowCount()):
            if not self._is_row_checked(row):
                continue
            node = self._nodes[row]
            is_root = (node.get("depth", 1) == 0)

            # Se già in PDM usa il doc_id esistente (no nuova creazione)
            existing_id = node.get("_existing_id")
            code_item = self.table.item(row, _COL_CODE)
            proposed_code = code_item.text() if code_item else ""

            # Ricalcola se ancora "—" (può accadere se la tabella non era visibile durante il populate)
            if proposed_code in ("—", "") and not existing_id:
                self._update_row_code(row)
                proposed_code = code_item.text() if code_item else ""

            if not proposed_code or proposed_code.startswith("ERR") or proposed_code == "—":
                if not existing_id:
                    if is_root:
                        QMessageBox.warning(
                            self, "Codice radice mancante",
                            "Impossibile generare il codice per l'assieme radice.\n"
                            "Verificare che la macchina sia selezionata nella prima riga."
                        )
                        return
                    continue  # salta componenti senza codice

            cmb_mac = self.table.cellWidget(row, _COL_MAC)
            cmb_grp = self.table.cellWidget(row, _COL_GRP)
            cmb_lvl = self.table.cellWidget(row, _COL_LVL)

            mid = cmb_mac.currentData() if cmb_mac else None
            gid = cmb_grp.currentData() if cmb_grp else None
            lvl_idx = cmb_lvl.currentIndex() if cmb_lvl else -1
            subtype = _LEVEL_OPTIONS[lvl_idx][1] if 0 <= lvl_idx < len(_LEVEL_OPTIONS) else "PRT"
            doc_level = _LEVEL_OPTIONS[lvl_idx][2] if 0 <= lvl_idx < len(_LEVEL_OPTIONS) else 2

            rows_to_import.append({
                "row":          row,
                "node":         node,
                "existing_id":  existing_id,
                "proposed_code": proposed_code,
                "machine_id":   mid,
                "group_id":     gid,
                "subtype":      subtype,
                "doc_level":    doc_level,
            })

        if not rows_to_import:
            QMessageBox.warning(
                self, "Nessuna riga",
                "Nessun componente selezionato con codice assegnato.\n"
                "Selezionare almeno un componente e assegnare la codifica."
            )
            return

        # Chiedi workspace di destinazione
        from config import load_local_config
        cfg = load_local_config()
        ws_root_str = cfg.get("sw_workspace", "")
        if not ws_root_str:
            QMessageBox.warning(
                self, "Workspace non configurata",
                "Configurare la workspace SolidWorks in Strumenti → Configurazione SolidWorks."
            )
            return
        ws_root = Path(ws_root_str)

        # 2) Genera codici reali e crea documenti nel DB
        progress = QProgressDialog(
            "Creazione documenti nel DB…", "Annulla", 0, len(rows_to_import) * 3, self
        )
        progress.setWindowTitle("Importazione ASM")
        progress.setMinimumDuration(0)
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setValue(0)

        step = 0
        # Mappa path_originale_upper → (new_code, doc_id, ext)
        path_to_code: dict[str, tuple[str, int, str]] = {}
        # Mappa stem_upper → new_code (per Pack and Go rename)
        stem_to_code: dict[str, str] = {}

        created = 0
        errors  = []

        for item in rows_to_import:
            if progress.wasCanceled():
                break
            step += 1
            progress.setValue(step)

            node       = item["node"]
            existing_id = item["existing_id"]
            subtype    = item["subtype"]
            mid        = item["machine_id"]
            gid        = item["group_id"]
            doc_level  = item["doc_level"]
            orig_path  = node["path"]
            ext        = Path(orig_path).suffix if orig_path else ".SLDPRT"
            from config import SW_EXTENSIONS
            doc_type_from_ext = SW_EXTENSIONS.get(ext.upper(), "Parte")

            if existing_id:
                # Usa documento già esistente
                existing_doc = self.db.fetchone(
                    "SELECT code, file_ext FROM documents WHERE id=?", (existing_id,)
                )
                if existing_doc:
                    code = existing_doc["code"]
                    real_ext = existing_doc.get("file_ext") or ext
                    path_to_code[orig_path.upper()] = (code, existing_id, real_ext)
                    stem_to_code[Path(orig_path).stem.upper()] = code
                continue

            # Genera codice reale
            try:
                if subtype == "ASM_MACH":
                    code = self.coding.next_code_liv0(mid)
                    doc_type = "Assieme"
                    gid = None
                elif subtype == "ASM_GRP":
                    code = self.coding.next_code_liv1(mid, gid)
                    doc_type = "Assieme"
                elif subtype == "ASM_SUB":
                    code = self.coding.next_code_liv2_subgroup(mid, gid)
                    doc_type = "Assieme"
                else:
                    code = self.coding.next_code_liv2_part(mid, gid)
                    doc_type = "Parte"

                # Il tipo reale è dettato dall'estensione del file
                if ext.upper() == ".SLDASM":
                    doc_type = "Assieme"
                elif ext.upper() == ".SLDDRW":
                    doc_type = "Disegno"
                else:
                    doc_type = doc_type  # già corretto

                stem = Path(orig_path).stem if orig_path else node["name"]
                doc_id = self.files.create_document(
                    code=code, revision="00", doc_type=doc_type,
                    title=stem,
                    machine_id=mid, group_id=gid, doc_level=doc_level,
                )
                path_to_code[orig_path.upper()] = (code, doc_id, ext)
                stem_to_code[Path(orig_path).stem.upper()] = code
                created += 1

                # Aggiorna la tabella con il codice reale
                code_item = self.table.item(item["row"], _COL_CODE)
                if code_item:
                    code_item.setText(code)

            except Exception as e:
                errors.append(f"{node['name']}: {e}")
                logging.error("Errore creazione doc %s: %s", node["name"], e, exc_info=True)

        if progress.wasCanceled():
            QMessageBox.warning(self, "Annullato", "Importazione annullata dall'utente.")
            return

        # 3) Pack and Go: copia workspace con rinomina e aggiornamento riferimenti
        progress.setLabelText("Copia struttura in workspace (Pack and Go)…")
        pgo_ok    = False
        pgo_error = ""
        pgo_count = 0

        try:
            pgo_count = self._run_pack_and_go(ws_root, stem_to_code)
            pgo_ok = True
            logging.info("Pack and Go OK: %d file copiati", pgo_count)
        except Exception as e:
            pgo_error = str(e)
            logging.error("Pack and Go fallito: %s", e, exc_info=True)

        # Fallback: se Pack and Go non disponibile, copia diretta (senza aggiorn. riferimenti)
        if not pgo_ok:
            progress.setLabelText("Copia diretta file in workspace (fallback)…")
            fallback_count = self._fallback_copy_files(ws_root, path_to_code)
            logging.info("Fallback copia diretta: %d file copiati", fallback_count)
            pgo_count = fallback_count
            # Anche con fallback procediamo ad archiviare ciò che c'è in workspace

        step = len(rows_to_import) * 2
        progress.setValue(step)

        # 3b) Fix riferimenti interni ASM (safety net: funziona sia dopo Pack&Go riuscito
        #     che dopo fallback; se i riferimenti sono già corretti non modifica nulla)
        progress.setLabelText("Aggiornamento riferimenti interni ASM…")
        fixed_refs = self._fix_asm_references(ws_root, path_to_code)
        if fixed_refs:
            logging.info("Riferimenti interni aggiornati in %d assieme/i", fixed_refs)

        # 4) Archivia i file presenti nella workspace (indipendente da Pack and Go)
        progress.setLabelText("Archiviazione file nel PDM…")
        archived = 0

        for k, (code, doc_id, ext) in path_to_code.items():
            if progress.wasCanceled():
                break
            ws_file = ws_root / (code + ext)
            if not ws_file.exists():
                logging.warning("File workspace non trovato, skip archiviazione: %s", ws_file)
                continue
            try:
                arch_dir = self.sp.archive_path(code, "00")
                arch_dir.mkdir(parents=True, exist_ok=True)
                arch_file = arch_dir / ws_file.name
                shutil.copy2(str(ws_file), str(arch_file))
                rel_path = str(arch_file.relative_to(self.sp.root))
                self.db.execute(
                    """UPDATE documents
                       SET file_name=?, file_ext=?, archive_path=?,
                           modified_at=datetime('now')
                       WHERE id=?""",
                    (arch_file.name, ext, rel_path, doc_id),
                )
                archived += 1
            except Exception as e:
                errors.append(f"Archiviazione {code}: {e}")
                logging.error("Archiviazione %s: %s", code, e)

        # 5) Collega BOM (relazioni padre-figlio)
        self._link_bom(path_to_code)

        step = len(rows_to_import) * 3
        progress.setValue(step)
        progress.close()

        # 6) Riepilogo
        ref_note = "" if pgo_ok else "\n⚠ Pack and Go non disponibile: file copiati senza aggiornamento riferimenti interni ASM."
        msg = (
            f"Importazione completata.\n\n"
            f"Documenti creati nel DB:    {created}\n"
            f"File copiati in workspace:  {pgo_count}\n"
            f"File archiviati nel PDM:    {archived}\n"
            f"\nWorkspace:  {ws_root}"
            f"{ref_note}"
        )
        if pgo_error and not pgo_ok:
            msg += f"\n\nDettaglio Pack and Go: {pgo_error}"
        if errors:
            msg += f"\n\nErrori ({len(errors)}):\n" + "\n".join(errors[:10])

        QMessageBox.information(self, "Importazione completata", msg)
        self.accept()

    # ------------------------------------------------------------------
    def _fallback_copy_files(self, ws_root: Path,
                              path_to_code: dict[str, tuple]) -> int:
        """
        Fallback se Pack and Go non è disponibile:
        copia ogni file originale nella workspace rinominandolo col codice PDM.
        Per il documento attivo in SolidWorks (ASM principale, che può essere
        locked in lettura), usa SaveAs3(..., 0, 2) = salva copia senza cambiare
        il percorso attivo in SW.
        I riferimenti interni dell'ASM NON vengono aggiornati.
        """
        ws_root.mkdir(parents=True, exist_ok=True)
        count = 0

        # Ottieni istanza SW e percorso documento attivo per il fallback SaveAs3
        sw_app          = None
        sw_active_path  = ""
        sw_active_doc   = None
        try:
            import win32com.client as win32
            sw_app = win32.GetActiveObject("SldWorks.Application")
            sw_doc = sw_app.ActiveDoc
            if sw_doc:
                sw_active_path = (sw_doc.GetPathName() or "").upper()
                sw_active_doc  = sw_doc
        except Exception:
            pass

        for orig_path_upper, (code, doc_id, ext) in path_to_code.items():
            # Recupera il percorso originale dal nodo
            orig_path = None
            for node in self._nodes:
                if node["path"] and node["path"].upper() == orig_path_upper:
                    orig_path = node["path"]
                    break
            if not orig_path:
                continue

            src  = Path(orig_path)
            dest = ws_root / (code + ext)
            copied = False

            # Tentativo 1: copia diretta con shutil
            try:
                shutil.copy2(str(src), str(dest))
                copied = True
                logging.info("Fallback copy2: %s → %s", src.name, dest.name)
            except Exception as e:
                logging.warning("copy2 fallita per %s: %s", src.name, e)

            # Tentativo 2: se è il documento attivo in SW, usa SaveAs3 (salva copia)
            if not copied and sw_active_doc and orig_path_upper == sw_active_path:
                try:
                    sw_active_doc.SaveAs3(str(dest).replace("/", "\\"), 0, 2)
                    if dest.exists():
                        copied = True
                        logging.info("Fallback SaveAs3 OK: %s", dest.name)
                    else:
                        logging.warning("SaveAs3 non ha creato il file: %s", dest)
                except Exception as e:
                    logging.warning("SaveAs3 fallita per %s: %s", src.name, e)

            if not src.exists() and not copied:
                logging.warning("Fallback: file sorgente non trovato: %s", src)
                continue

            if copied:
                count += 1
            else:
                logging.warning("Fallback: impossibile copiare %s", src.name)

        return count

    # ------------------------------------------------------------------
    def _run_pack_and_go(self, ws_root: Path, stem_to_code: dict[str, str]) -> int:
        """
        Usa SolidWorks Pack and Go COM API per copiare l'assieme nella workspace
        rinominando i file con il codice PDM e aggiornando i riferimenti interni.
        Ritorna il numero di file copiati.
        """
        import win32com.client as win32
        import pythoncom

        try:
            sw = win32.GetActiveObject("SldWorks.Application")
        except Exception:
            raise RuntimeError(
                "SolidWorks non è in esecuzione.\n"
                "Aprire l'assieme in SolidWorks per poter usare Pack and Go."
            )

        model = sw.ActiveDoc
        if model is None:
            raise RuntimeError("Nessun documento attivo in SolidWorks.")

        pgo = None
        try:
            pgo = model.Extension.GetPackAndGo()
        except Exception as e:
            raise RuntimeError(f"Pack and Go non disponibile: {e}")

        if pgo is None:
            raise RuntimeError("GetPackAndGo() ha restituito None.")

        ws_root.mkdir(parents=True, exist_ok=True)

        try:
            pgo.IncludeDrawings          = False
            pgo.IncludeSimulationResults = False
            pgo.IncludeToolboxComponents = False
        except Exception:
            pass

        # STEP 1: leggi nomi PRIMA di impostare la destinazione (ordine critico)
        try:
            raw_names = pgo.GetDocumentNames()
            file_names = None
            if isinstance(raw_names, tuple):
                for item in raw_names:
                    if item is not None and hasattr(item, '__iter__') and not isinstance(item, (str, int, float, bool)):
                        candidates = list(item)
                        if candidates and isinstance(candidates[0], str):
                            file_names = candidates
                            break
                if file_names is None:
                    file_names = [str(x) for x in raw_names if isinstance(x, str)]
            elif raw_names is not None:
                file_names = list(raw_names)
            if not file_names:
                file_names = []
        except Exception as e:
            raise RuntimeError(f"GetDocumentNames fallito: {e}")

        if not file_names:
            raise RuntimeError("Pack and Go: nessun file trovato nel pack.")

        # STEP 2: costruisci nuovi percorsi (backslash obbligatorio su Windows)
        new_names = []
        for fname in file_names:
            stem = Path(fname).stem.upper()
            new_code = stem_to_code.get(stem)
            if new_code:
                new_name = str(ws_root / (new_code + Path(fname).suffix)).replace("/", "\\")
            else:
                new_name = str(ws_root / Path(fname).name).replace("/", "\\")
            new_names.append(new_name)

        # STEP 3: cartella radice (safety net per file non mappati esplicitamente)
        try:
            pgo.SetRootSaveDirectory(str(ws_root).replace("/", "\\"))
        except Exception as e:
            raise RuntimeError(f"SetRootSaveDirectory fallito: {e}")

        # STEP 4: imposta percorsi con array COM tipizzati (VT_ARRAY|VT_BSTR)
        # win32com non converte automaticamente le liste Python in SafeArray di BSTR;
        # senza questa conversione SetDocumentSavePaths copia i file ma NON aggiorna
        # i riferimenti interni degli ASM.
        try:
            fn_var = win32.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_BSTR, file_names)
            nn_var = win32.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_BSTR, new_names)
            ok = pgo.SetDocumentSavePaths(fn_var, nn_var)
            logging.info("SetDocumentSavePaths result: %s", ok)
        except Exception as e:
            raise RuntimeError(f"SetDocumentSavePaths fallito: {e}")

        # STEP 5: esegui
        try:
            save_result = pgo.Save()
            if isinstance(save_result, tuple):
                save_status = save_result[0]
            else:
                save_status = save_result if isinstance(save_result, int) else 0
            if save_status not in (None, 0):
                logging.warning("Pack and Go Save() status: %s", save_status)
        except Exception as e:
            raise RuntimeError(f"Pack and Go Save() fallito: {e}")

        count = sum(1 for p in new_names if Path(p).exists())
        if count == 0:
            raise RuntimeError(
                "Pack and Go Save() non ha creato alcun file nella workspace.\n"
                "Verificare che l'assieme sia aperto e completamente risolto in SolidWorks."
            )
        return count

    # ------------------------------------------------------------------
    def _fix_asm_references(self, ws_root: Path,
                             path_to_code: dict[str, tuple]) -> int:
        """
        Post-processing di sicurezza: apre ogni SLDASM copiato nella workspace
        e sostituisce i riferimenti ancora puntati ai file originali (non codificati)
        con i nuovi percorsi codificati tramite ReplaceReferencedDocument.
        Viene chiamato sempre dopo Pack&Go o fallback per garantire riferimenti corretti.
        Ritorna il numero di assieme modificati.
        """
        # Mappa orig_path_upper → new_full_path
        orig_to_new: dict[str, str] = {}
        for orig_upper, (code, doc_id, ext) in path_to_code.items():
            new_path = str(ws_root / (code + ext)).replace("/", "\\")
            orig_to_new[orig_upper] = new_path

        # Mappa per recuperare path originale con case corretto dai nodi
        orig_case: dict[str, str] = {}
        for node in self._nodes:
            if node.get("path"):
                orig_case[node["path"].upper()] = node["path"]

        try:
            import win32com.client as win32
            import pythoncom
            sw = win32.GetActiveObject("SldWorks.Application")
        except Exception:
            return 0

        modified = 0
        for orig_upper, (code, doc_id, ext) in path_to_code.items():
            if ext.upper() != ".SLDASM":
                continue
            asm_ws = ws_root / (code + ext)
            if not asm_ws.exists():
                continue

            asm_str = str(asm_ws).replace("/", "\\")
            opened_here = False
            asm_model = None
            try:
                # Se già aperto in SW (es. è la radice ancora attiva) non riaprire
                try:
                    asm_model = sw.GetOpenDocumentByName(asm_str)
                except Exception:
                    asm_model = None

                if asm_model is None:
                    errors_v   = win32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
                    warnings_v = win32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
                    # swDocASSEMBLY=2, swOpenDocOptions_Silent=1
                    asm_model = sw.OpenDoc6(asm_str, 2, 1, "", errors_v, warnings_v)
                    if asm_model is None:
                        continue
                    opened_here = True

                # Scorri tutti i componenti
                try:
                    raw = asm_model.GetComponents(False)
                    comps = list(raw) if raw is not None else []
                except Exception:
                    comps = []

                changed = False
                seen: set = set()
                for comp in comps:
                    try:
                        v = comp.GetPathName
                        if callable(v):
                            v = v()
                        fp = str(v or "").strip()
                        if not fp:
                            continue
                        fp_upper = fp.upper()
                        if fp_upper in seen:
                            continue
                        seen.add(fp_upper)
                        if fp_upper in orig_to_new:
                            new_fp = orig_to_new[fp_upper]
                            # Sostituisci solo se punta ancora al file originale
                            # (path diverso dalla destinazione) e il nuovo file esiste
                            if fp_upper != new_fp.upper() and Path(new_fp).exists():
                                ok = asm_model.Extension.ReplaceReferencedDocument(fp, new_fp)
                                if ok:
                                    changed = True
                                    logging.info("ReplaceRef OK: %s → %s",
                                                 Path(fp).name, Path(new_fp).name)
                                else:
                                    logging.warning("ReplaceRef fallito: %s → %s",
                                                    Path(fp).name, Path(new_fp).name)
                    except Exception as e_comp:
                        logging.warning("Errore componente in fix_refs: %s", e_comp)

                if changed:
                    try:
                        asm_model.Save3(1, 0, 0)
                        modified += 1
                        logging.info("ASM salvato dopo fix riferimenti: %s", asm_ws.name)
                    except Exception as e_save:
                        logging.warning("Salvataggio ASM fallito: %s", e_save)

            except Exception as e_asm:
                logging.warning("Fix riferimenti ASM %s: %s", asm_ws.name, e_asm)
            finally:
                if opened_here and asm_model:
                    try:
                        title = asm_model.GetTitle()
                        if callable(title):
                            title = title()
                        if title:
                            sw.CloseDoc(title)
                    except Exception:
                        pass

        return modified

    # ------------------------------------------------------------------
    def _link_bom(self, path_to_code: dict[str, tuple[str, int, str]]):
        """Aggiunge le relazioni BOM padre-figlio per i componenti importati."""
        for node in self._nodes:
            if not node.get("path") or not node.get("parent_path"):
                continue
            child_key  = node["path"].upper()
            parent_key = node["parent_path"].upper()
            child_entry  = path_to_code.get(child_key)
            parent_entry = path_to_code.get(parent_key)
            if not child_entry or not parent_entry:
                continue
            _, child_id, _  = child_entry
            _, parent_id, _ = parent_entry
            try:
                self.asm_mgr.add_component(
                    parent_id, child_id,
                    float(node.get("quantity", 1))
                )
            except Exception as e:
                logging.warning("Link BOM %s → %s: %s", parent_id, child_id, e)

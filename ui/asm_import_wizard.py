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
        # Refresh unico: anteprime sequenziali + evidenziazione righe attive
        self._refresh_preview_codes()

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

        self._apply_row_checked_style(row, self._is_row_checked(row))

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

    def _apply_row_checked_style(self, row: int, checked: bool):
        """Evidenzia le righe attive (checkbox selezionata)."""
        bg_active = QColor("#313244")
        bg_base = QColor("#181825") if row % 2 == 0 else QColor("#1a1a2e")
        bg = bg_active if checked else bg_base

        for col in range(self.table.columnCount()):
            item = self.table.item(row, col)
            if item:
                item.setBackground(bg)

        for col in (_COL_INC, _COL_MAC, _COL_GRP, _COL_LVL):
            w = self.table.cellWidget(row, col)
            if w:
                w.setStyleSheet(f"background-color: {bg.name()};")

    def _peek_counter_value(self, counter_type: str, machine_id, group_id) -> int:
        row = self.db.fetchone(
            """SELECT last_value
               FROM hierarchical_counters
               WHERE counter_type=? AND machine_id IS ? AND group_id IS ?""",
            (counter_type, machine_id, group_id),
        )
        return int(row["last_value"]) if row else 0

    def _refresh_preview_codes(self):
        """
        Ricalcola tutte le preview in modo sequenziale simulato.
        Le righe selezionate mostrano codici progressivi univoci.
        """
        counter_state: dict[tuple[str, int, Optional[int]], int] = {}
        machine_code_cache: dict[int, str] = {}
        group_code_cache: dict[int, str] = {}

        for row in range(self.table.rowCount()):
            item = self.table.item(row, _COL_CODE)
            if not item:
                continue

            checked = self._is_row_checked(row)
            self._apply_row_checked_style(row, checked)

            node = self._nodes[row] if row < len(self._nodes) else None
            if node and node.get("_existing_id"):
                existing = self.db.fetchone(
                    "SELECT code FROM documents WHERE id=?", (node["_existing_id"],)
                )
                item.setText(existing["code"] if existing else "gia nel PDM")
                item.setForeground(QColor("#a6e3a1"))
                continue

            if not checked:
                item.setText("—")
                item.setForeground(QColor("#6c7086"))
                continue

            cmb_mac = self.table.cellWidget(row, _COL_MAC)
            cmb_grp = self.table.cellWidget(row, _COL_GRP)
            cmb_lvl = self.table.cellWidget(row, _COL_LVL)
            if not (cmb_mac and cmb_grp and cmb_lvl):
                item.setText("—")
                item.setForeground(QColor("#6c7086"))
                continue

            mid = cmb_mac.currentData()
            gid = cmb_grp.currentData()
            lvl_idx = cmb_lvl.currentIndex()
            subtype = _LEVEL_OPTIONS[lvl_idx][1] if 0 <= lvl_idx < len(_LEVEL_OPTIONS) else "PRT"
            key = self._counter_key_for_row(subtype, mid, gid)

            if not mid or not key:
                item.setText("ERR: selezione")
                item.setForeground(QColor("#f38ba8"))
                continue

            counter_type, key_mid, key_gid = key
            state_key = (counter_type, key_mid, key_gid)
            cur = counter_state.get(state_key)
            if cur is None:
                cur = self._peek_counter_value(counter_type, key_mid, key_gid)

            if counter_type == "SUBGROUP":
                nxt = 9999 if cur == 0 else cur - 1
            else:
                nxt = cur + 1
            counter_state[state_key] = nxt

            try:
                if key_mid not in machine_code_cache:
                    m = self.coding.get_machine(key_mid)
                    if not m:
                        raise ValueError("macchina non trovata")
                    machine_code_cache[key_mid] = m["code"]
                m_code = machine_code_cache[key_mid]

                if subtype == "ASM_MACH":
                    code_text = f"{m_code}_V{nxt:03d}"
                else:
                    if key_gid is None:
                        raise ValueError("gruppo mancante")
                    if key_gid not in group_code_cache:
                        g = self.coding.get_group(key_gid)
                        if not g:
                            raise ValueError("gruppo non trovato")
                        group_code_cache[key_gid] = g["code"]
                    g_code = group_code_cache[key_gid]

                    if subtype == "ASM_GRP":
                        code_text = f"{m_code}_{g_code}-V{nxt:03d}"
                    else:
                        code_text = f"{m_code}_{g_code}-{nxt:04d}"

                        if subtype == "PRT" and nxt > 8999:
                            raise ValueError("PART esaurito")
                        if subtype == "ASM_SUB" and nxt < 9000:
                            raise ValueError("SUBGROUP esaurito")

                item.setText(code_text)
                item.setForeground(QColor("#89b4fa"))
            except Exception as e:
                item.setText(f"ERR: {e}")
                item.setForeground(QColor("#f38ba8"))

    # ------------------------------------------------------------------
    # Segnali riga
    # ------------------------------------------------------------------
    def _on_check_changed(self, row: int, state: int):
        if row < len(self._nodes) and self._nodes[row].get("depth", 1) == 0:
            self._set_row_checked(row, True, emit=False)
            self._refresh_preview_codes()
            return

        checked = (state == Qt.CheckState.Checked.value)
        for col in (_COL_MAC, _COL_GRP, _COL_LVL):
            w = self.table.cellWidget(row, col)
            if w:
                w.setEnabled(checked)
        self._apply_row_checked_style(row, checked)
        self._refresh_preview_codes()

    def _on_row_machine_changed(self, row: int):
        cmb_mac = self.table.cellWidget(row, _COL_MAC)
        cmb_grp = self.table.cellWidget(row, _COL_GRP)
        if cmb_mac and cmb_grp:
            self._fill_group_combo(cmb_grp, cmb_mac.currentData())
        self._refresh_preview_codes()

    def _update_row_code(self, row: int):
        # Le preview dipendono dai contatori simulati e dall'ordine righe:
        # aggiorniamo sempre l'intera tabella.
        self._refresh_preview_codes()

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

        self._refresh_preview_codes()

    def _select_all(self):
        for row in range(self.table.rowCount()):
            self._set_row_checked(row, True, emit=False)
        self._refresh_preview_codes()

    def _deselect_all(self):
        for row in range(self.table.rowCount()):
            self._set_row_checked(row, False, emit=False)
        self._refresh_preview_codes()

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def _is_row_checked(self, row: int) -> bool:
        if 0 <= row < len(self._nodes) and self._nodes[row].get("depth", 1) == 0:
            return True
        w = self.table.cellWidget(row, _COL_INC)
        if w:
            chk = w.findChild(QCheckBox)
            if chk:
                return chk.isChecked()
        return False

    def _set_row_checked(self, row: int, checked: bool, emit: bool = True):
        if 0 <= row < len(self._nodes) and self._nodes[row].get("depth", 1) == 0:
            checked = True
        w = self.table.cellWidget(row, _COL_INC)
        if w:
            chk = w.findChild(QCheckBox)
            if chk:
                if emit:
                    chk.setChecked(checked)
                else:
                    chk.blockSignals(True)
                    chk.setChecked(checked)
                    chk.blockSignals(False)
                    for col in (_COL_MAC, _COL_GRP, _COL_LVL):
                        cw = self.table.cellWidget(row, col)
                        if cw:
                            cw.setEnabled(checked)
                    self._apply_row_checked_style(row, checked)

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
    def _counter_key_for_row(self, subtype: str, machine_id, group_id):
        if not machine_id:
            return None
        if subtype == "ASM_MACH":
            return ("VERSION", machine_id, None)
        if group_id is None:
            return None
        if subtype == "ASM_GRP":
            return ("VERSION", machine_id, group_id)
        if subtype == "ASM_SUB":
            return ("SUBGROUP", machine_id, group_id)
        return ("PART", machine_id, group_id)

    def _snapshot_counters(self, rows_to_import: list[dict]) -> dict:
        """Snapshot dei contatori usati dall'import corrente."""
        keys: set = set()
        for item in rows_to_import:
            if item.get("existing_id"):
                continue
            key = self._counter_key_for_row(
                item.get("subtype", ""),
                item.get("machine_id"),
                item.get("group_id"),
            )
            if key:
                keys.add(key)

        snap: dict = {}
        for counter_type, machine_id, group_id in keys:
            row = self.db.fetchone(
                """SELECT id, last_value
                   FROM hierarchical_counters
                   WHERE counter_type=? AND machine_id IS ? AND group_id IS ?""",
                (counter_type, machine_id, group_id),
            )
            snap[(counter_type, machine_id, group_id)] = {
                "exists": bool(row),
                "value": int(row["last_value"]) if row else 0,
            }
        return snap

    def _rollback_failed_import(self,
                                created_doc_ids: list[int],
                                counter_snapshot: dict,
                                ws_cleanup: list[Path],
                                archive_cleanup: list[Path]) -> list[str]:
        """Rollback best-effort di file, documenti e contatori."""
        notes: list[str] = []
        rb_errors: list[str] = []

        removed_ws = 0
        for fp in ws_cleanup:
            try:
                if fp.exists():
                    fp.unlink()
                    removed_ws += 1
            except Exception as e:
                rb_errors.append(f"workspace {fp.name}: {e}")
        if removed_ws:
            notes.append(f"- File workspace rimossi: {removed_ws}")

        removed_arch = 0
        for fp in archive_cleanup:
            try:
                if fp.exists():
                    fp.unlink()
                    removed_arch += 1
            except Exception as e:
                rb_errors.append(f"archivio {fp.name}: {e}")
        if removed_arch:
            notes.append(f"- File archivio rimossi: {removed_arch}")

        if created_doc_ids:
            placeholders = ",".join("?" for _ in created_doc_ids)
            p1 = tuple(created_doc_ids)
            p2 = tuple(created_doc_ids) + tuple(created_doc_ids)

            try:
                self.db.execute(
                    f"""DELETE FROM asm_components
                        WHERE parent_id IN ({placeholders})
                           OR child_id  IN ({placeholders})""",
                    p2,
                )
            except Exception as e:
                rb_errors.append(f"asm_components: {e}")

            for table in (
                "document_properties",
                "document_versions",
                "checkout_log",
                "workspace_files",
                "workflow_history",
            ):
                try:
                    self.db.execute(
                        f"DELETE FROM {table} WHERE document_id IN ({placeholders})",
                        p1,
                    )
                except Exception as e:
                    rb_errors.append(f"{table}: {e}")

            try:
                self.db.execute(
                    f"DELETE FROM documents WHERE id IN ({placeholders})",
                    p1,
                )
                notes.append(f"- Documenti DB eliminati: {len(created_doc_ids)}")
            except Exception as e:
                rb_errors.append(f"documents: {e}")

        restored = 0
        for (counter_type, machine_id, group_id), row in counter_snapshot.items():
            try:
                if row.get("exists"):
                    self.coding.reset_counter(
                        counter_type, machine_id, group_id, int(row.get("value", 0))
                    )
                else:
                    self.db.execute(
                        """DELETE FROM hierarchical_counters
                           WHERE counter_type=? AND machine_id IS ? AND group_id IS ?""",
                        (counter_type, machine_id, group_id),
                    )
                restored += 1
            except Exception as e:
                rb_errors.append(f"counter {counter_type}/{machine_id}/{group_id}: {e}")
        if restored:
            notes.append(f"- Contatori ripristinati: {restored}")

        if rb_errors:
            notes.append("- Errori rollback: " + "; ".join(rb_errors[:6]))
        return notes

    def _validate_asm_references(self, ws_root: Path,
                                 path_to_code: dict[str, tuple]) -> list[str]:
        """
        Verifica hard: nessun ASM copiato deve referenziare i path originali
        mappati in questa importazione.
        """
        orig_to_new: dict[str, str] = {}
        for orig_upper, (code, _doc_id, ext) in path_to_code.items():
            orig_to_new[orig_upper] = str(ws_root / (code + ext)).replace("/", "\\")

        # Validazione hard solo per i componenti rinominati (codificati).
        required_refs: set[str] = set()
        for orig_upper, new_fp in orig_to_new.items():
            old_name = Path(orig_upper).name.upper()
            new_name = Path(new_fp).name.upper()
            if old_name != new_name:
                required_refs.add(orig_upper)

        try:
            import win32com.client as win32
            import pythoncom
            sw = win32.GetActiveObject("SldWorks.Application")
        except Exception as e:
            raise RuntimeError(f"Validazione riferimenti non disponibile: {e}")

        unresolved: list[str] = []
        seen_unresolved: set = set()

        for _orig_upper, (code, _doc_id, ext) in path_to_code.items():
            if ext.upper() != ".SLDASM":
                continue
            asm_ws = ws_root / (code + ext)
            if not asm_ws.exists():
                unresolved.append(f"{asm_ws.name}: file ASM non presente in workspace")
                continue

            asm_str = str(asm_ws).replace("/", "\\")
            asm_model = None
            opened_here = False
            try:
                try:
                    asm_model = sw.GetOpenDocumentByName(asm_str)
                except Exception:
                    asm_model = None

                if asm_model is None:
                    errors_v = win32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
                    warnings_v = win32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
                    asm_model = sw.OpenDoc6(asm_str, 2, 1, "", errors_v, warnings_v)
                    if asm_model is None:
                        unresolved.append(f"{asm_ws.name}: impossibile aprire ASM per validazione")
                        continue
                    opened_here = True

                try:
                    raw = asm_model.GetComponents(False)
                    comps = list(raw) if raw is not None else []
                except Exception:
                    comps = []

                for comp in comps:
                    try:
                        fp = comp.GetPathName
                        if callable(fp):
                            fp = fp()
                        fp = str(fp or "").strip()
                    except Exception:
                        fp = ""
                    if not fp:
                        continue
                    fp_upper = fp.upper()
                    if fp_upper in required_refs:
                        key = f"{asm_ws.name}|{fp_upper}"
                        if key in seen_unresolved:
                            continue
                        seen_unresolved.add(key)
                        unresolved.append(
                            f"{asm_ws.name}: riferimento non aggiornato -> {Path(fp).name}"
                        )
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

        return unresolved

    def _import(self):
        """Orchestra l'importazione in modalita fail-safe (rollback completo)."""
        # 1) Raccogli righe selezionate con codice proposto
        rows_to_import: list[dict] = []
        for row in range(self.table.rowCount()):
            if not self._is_row_checked(row):
                continue
            node = self._nodes[row]
            is_root = (node.get("depth", 1) == 0)

            existing_id = node.get("_existing_id")
            code_item = self.table.item(row, _COL_CODE)
            proposed_code = code_item.text() if code_item else ""

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
                    continue

            cmb_mac = self.table.cellWidget(row, _COL_MAC)
            cmb_grp = self.table.cellWidget(row, _COL_GRP)
            cmb_lvl = self.table.cellWidget(row, _COL_LVL)

            mid = cmb_mac.currentData() if cmb_mac else None
            gid = cmb_grp.currentData() if cmb_grp else None
            lvl_idx = cmb_lvl.currentIndex() if cmb_lvl else -1
            subtype = _LEVEL_OPTIONS[lvl_idx][1] if 0 <= lvl_idx < len(_LEVEL_OPTIONS) else "PRT"
            doc_level = _LEVEL_OPTIONS[lvl_idx][2] if 0 <= lvl_idx < len(_LEVEL_OPTIONS) else 2

            rows_to_import.append({
                "row": row,
                "node": node,
                "existing_id": existing_id,
                "proposed_code": proposed_code,
                "machine_id": mid,
                "group_id": gid,
                "subtype": subtype,
                "doc_level": doc_level,
            })

        if not rows_to_import:
            QMessageBox.warning(
                self, "Nessuna riga",
                "Nessun componente selezionato con codice assegnato.\n"
                "Selezionare almeno un componente e assegnare la codifica."
            )
            return

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

        progress = QProgressDialog(
            "Creazione documenti nel DB...", "Annulla", 0, len(rows_to_import) * 4, self
        )
        progress.setWindowTitle("Importazione ASM")
        progress.setMinimumDuration(0)
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setValue(0)

        counter_snapshot = self._snapshot_counters(rows_to_import)
        created_doc_ids: list[int] = []
        created_doc_id_set: set[int] = set()
        ws_cleanup: list[Path] = []
        archive_cleanup: list[Path] = []

        step = 0
        path_to_code: dict[str, tuple[str, int, str]] = {}
        stem_to_code: dict[str, str] = {}

        created = 0
        errors = []
        pgo_ok = False
        pgo_error = ""
        pgo_count = 0
        fixed_refs = 0
        archived = 0
        bom_reused_uncoded = 0
        bom_created_uncoded = 0

        try:
            for item in rows_to_import:
                if progress.wasCanceled():
                    raise RuntimeError("Importazione annullata dall'utente.")
                step += 1
                progress.setValue(step)

                node = item["node"]
                existing_id = item["existing_id"]
                subtype = item["subtype"]
                mid = item["machine_id"]
                gid = item["group_id"]
                doc_level = item["doc_level"]
                orig_path = node["path"]
                ext = Path(orig_path).suffix if orig_path else ".SLDPRT"

                if existing_id:
                    existing_doc = self.db.fetchone(
                        "SELECT code, file_ext FROM documents WHERE id=?", (existing_id,)
                    )
                    if existing_doc:
                        code = existing_doc["code"]
                        real_ext = existing_doc.get("file_ext") or ext
                        path_to_code[orig_path.upper()] = (code, existing_id, real_ext)
                        stem_to_code[Path(orig_path).stem.upper()] = code
                    continue

                if subtype == "ASM_MACH":
                    code = self.coding.next_code_liv0(mid)
                    doc_type = "Assieme"
                    gid = None
                elif subtype == "ASM_GRP":
                    code = self.coding.next_code_liv1(mid, gid)
                    doc_type = "Assieme"
                elif subtype == "ASM_SUB":
                    code = self.coding.next_code_liv2_part(mid, gid)
                    doc_type = "Assieme"
                else:
                    code = self.coding.next_code_liv2_subgroup(mid, gid)
                    doc_type = "Parte"

                if ext.upper() == ".SLDASM":
                    doc_type = "Assieme"
                elif ext.upper() == ".SLDDRW":
                    doc_type = "Disegno"

                stem = Path(orig_path).stem if orig_path else node["name"]
                doc_id = self.files.create_document(
                    code=code, revision="00", doc_type=doc_type,
                    title=stem,
                    machine_id=mid, group_id=gid, doc_level=doc_level,
                )
                created_doc_ids.append(doc_id)
                created_doc_id_set.add(doc_id)
                path_to_code[orig_path.upper()] = (code, doc_id, ext)
                stem_to_code[Path(orig_path).stem.upper()] = code
                created += 1

                code_item = self.table.item(item["row"], _COL_CODE)
                if code_item:
                    code_item.setText(code)

            if not path_to_code:
                raise RuntimeError("Nessun percorso valido da importare.")

            seen_ws: set = set()
            for _k, (code, _doc_id, ext) in path_to_code.items():
                fp = ws_root / (code + ext)
                key = str(fp).upper()
                if key in seen_ws:
                    continue
                seen_ws.add(key)
                if not fp.exists():
                    ws_cleanup.append(fp)

            progress.setLabelText("Copia struttura in workspace (Pack and Go)...")
            try:
                pgo_count = self._run_pack_and_go(ws_root, stem_to_code)
                pgo_ok = True
                logging.info("Pack and Go OK: %d file copiati", pgo_count)
            except Exception as e:
                pgo_error = str(e)
                logging.error("Pack and Go fallito: %s", e, exc_info=True)

            if not pgo_ok:
                progress.setLabelText("Copia diretta file in workspace (fallback)...")
                fallback_count, fallback_created = self._fallback_copy_files(ws_root, path_to_code)
                for fp in fallback_created:
                    if fp not in ws_cleanup:
                        ws_cleanup.append(fp)
                logging.info("Fallback copia diretta: %d file copiati", fallback_count)
                pgo_count = fallback_count

            step = len(rows_to_import) * 2
            progress.setValue(step)

            if pgo_count <= 0:
                raise RuntimeError("Nessun file copiato in workspace. Importazione interrotta.")

            # Completa mappatura documenti per l'intera struttura ASM:
            # anche i nodi non codificati vengono collegati in BOM.
            bom_reused_uncoded, bom_created_uncoded = self._ensure_bom_document_map(
                path_to_code, created_doc_ids, created_doc_id_set
            )

            progress.setLabelText("Aggiornamento riferimenti interni ASM...")
            fixed_refs = self._fix_asm_references(ws_root, path_to_code)
            if fixed_refs:
                logging.info("Riferimenti interni aggiornati in %d assieme/i", fixed_refs)

            progress.setLabelText("Validazione riferimenti ASM...")
            unresolved = self._validate_asm_references(ws_root, path_to_code)
            if unresolved:
                preview = "\n".join(unresolved[:12])
                raise RuntimeError(
                    "Riferimenti interni ASM non coerenti dopo copia/rinomina.\n"
                    f"{preview}"
                )

            progress.setLabelText("Archiviazione file nel PDM...")
            from core.checkout_manager import CheckoutManager
            co_mgr = CheckoutManager(self.db, self.sp, self.user)

            for _k, (code, doc_id, ext) in path_to_code.items():
                if progress.wasCanceled():
                    raise RuntimeError("Importazione annullata dall'utente.")
                if doc_id not in created_doc_id_set:
                    continue

                ws_file = ws_root / (code + ext)
                if not ws_file.exists():
                    raise RuntimeError(f"File workspace mancante: {ws_file.name}")

                # Traccia arch_file per cleanup rollback
                arch_dir = self.sp.archive_path(code, "00")
                arch_file = arch_dir / (code + ext)
                arch_preexisting = arch_file.exists()

                # Archivia, aggiorna DB, imposta checkout lock e permessi
                co_mgr.checkout_new_from_workspace(doc_id, ws_file)

                if not arch_preexisting:
                    archive_cleanup.append(arch_file)
                archived += 1

            self._link_bom(path_to_code)

        except Exception as e:
            logging.error("Importazione ASM fallita: %s", e, exc_info=True)
            rb_notes = self._rollback_failed_import(
                created_doc_ids, counter_snapshot, ws_cleanup, archive_cleanup
            )
            progress.close()
            msg = (
                "Importazione annullata: operazione non valida o incompleta.\n\n"
                f"Motivo:\n{e}\n\n"
                "Rollback eseguito."
            )
            if rb_notes:
                msg += "\n" + "\n".join(rb_notes)
            QMessageBox.critical(self, "Importazione annullata", msg)
            return

        step = len(rows_to_import) * 4
        progress.setValue(step)
        progress.close()

        ref_note = "" if pgo_ok else "\nNota: usato fallback copia diretta + fix riferimenti."
        msg = (
            f"Importazione completata.\n\n"
            f"Documenti creati nel DB:    {created}\n"
            f"Non codificati creati DB:   {bom_created_uncoded}\n"
            f"Non codificati riusati DB:  {bom_reused_uncoded}\n"
            f"File copiati in workspace:  {pgo_count}\n"
            f"ASM con fix riferimenti:    {fixed_refs}\n"
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
                              path_to_code: dict[str, tuple]) -> tuple[int, list[Path]]:
        """
        Fallback se Pack and Go non è disponibile:
        copia ogni file selezionato nella workspace rinominandolo col codice PDM
        e copia anche i file non selezionati con nome originale.
        Per il documento attivo in SolidWorks (ASM principale, che può essere
        locked in lettura), usa SaveAs3(..., 0, 2) = salva copia senza cambiare
        il percorso attivo in SW.
        I riferimenti interni dell'ASM NON vengono aggiornati.
        """
        ws_root.mkdir(parents=True, exist_ok=True)
        count = 0
        created_paths: list[Path] = []

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

        mapped_destinations: set[str] = {
            str(ws_root / (code + ext)).upper()
            for _orig, (code, _doc_id, ext) in path_to_code.items()
        }

        # 1) Copia dei file selezionati (rinominati con codice PDM)
        for orig_path_upper, (code, _doc_id, ext) in path_to_code.items():
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
            dest_preexists = dest.exists()
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
                if not dest_preexists and dest.exists():
                    created_paths.append(dest)
            else:
                logging.warning("Fallback: impossibile copiare %s", src.name)

        # 2) Copia file NON selezionati (nome originale, senza codifica)
        copied_unselected: set[str] = set()
        selected_keys = set(path_to_code.keys())
        for node in self._nodes:
            orig_path = node.get("path")
            if not orig_path:
                continue

            orig_upper = orig_path.upper()
            if orig_upper in selected_keys:
                continue

            src = Path(orig_path)
            if not src.exists():
                logging.warning("Fallback non selezionati: file sorgente non trovato: %s", src)
                continue

            dest = ws_root / src.name
            dest_upper = str(dest).upper()

            # Evita di sovrascrivere file codificati della selezione
            if dest_upper in mapped_destinations:
                logging.warning(
                    "Fallback non selezionati: skip %s (collisione con file codificato)",
                    src.name,
                )
                continue

            # Evita copie duplicate quando più nodi puntano allo stesso nome file
            if dest_upper in copied_unselected:
                continue

            try:
                if not dest.exists():
                    shutil.copy2(str(src), str(dest))
                    created_paths.append(dest)
                copied_unselected.add(dest_upper)
                count += 1
                logging.info("Fallback non selezionati copy2: %s → %s", src.name, dest.name)
            except Exception as e:
                logging.warning(
                    "Fallback non selezionati: impossibile copiare %s: %s",
                    src.name,
                    e,
                )

        return count, created_paths

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
        except Exception as e_noarg:
            try:
                status_v = win32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
                pgo = model.Extension.GetPackAndGo(status_v)
                logging.info("GetPackAndGo(status) OK: %s", status_v)
            except Exception as e_witharg:
                raise RuntimeError(
                    f"Pack and Go non disponibile: {e_witharg} "
                    f"(fallback no-arg: {e_noarg})"
                )

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
        Post-processing di sicurezza: sostituisce i riferimenti ancora puntati ai
        file originali con i nuovi percorsi codificati.
        Usa ISldWorks.ReplaceReferencedDocument sul documento ASM chiuso.
        Ritorna il numero di assieme modificati.
        """
        # Mappa orig_path_upper -> new_full_path
        orig_to_new: dict[str, str] = {}
        for orig_upper, (code, _doc_id, ext) in path_to_code.items():
            new_path = str(ws_root / (code + ext)).replace("/", "\\")
            orig_to_new[orig_upper] = new_path

        # Da aggiornare obbligatoriamente solo i riferimenti che cambiano nome file.
        # Per i non codificati (stesso basename) non forziamo il replace hard.
        orig_to_new_required: dict[str, str] = {}
        for orig_upper, new_fp in orig_to_new.items():
            old_name = Path(orig_upper).name.upper()
            new_name = Path(new_fp).name.upper()
            if old_name != new_name:
                orig_to_new_required[orig_upper] = new_fp

        # Mappa per recuperare path originale con case corretto
        orig_case: dict[str, str] = {}
        for node in self._nodes:
            if node.get("path"):
                orig_case[node["path"].upper()] = node["path"]

        try:
            import win32com.client as win32
            sw = win32.GetActiveObject("SldWorks.Application")
        except Exception:
            return 0

        modified = 0
        for _orig_upper, (code, _doc_id, ext) in path_to_code.items():
            if ext.upper() != ".SLDASM":
                continue
            asm_ws = ws_root / (code + ext)
            if not asm_ws.exists():
                continue

            asm_str = str(asm_ws).replace("/", "\\")
            try:
                # ReplaceReferencedDocument richiede il documento referencing chiuso
                try:
                    asm_model = sw.GetOpenDocumentByName(asm_str)
                except Exception:
                    asm_model = None
                if asm_model is not None:
                    try:
                        title = asm_model.GetTitle()
                        if callable(title):
                            title = title()
                        if title:
                            sw.CloseDoc(title)
                    except Exception as e_close:
                        logging.warning("CloseDoc prima di ReplaceRef (%s): %s", asm_ws.name, e_close)

                changed = False
                for orig_upper, new_fp in orig_to_new_required.items():
                    if orig_upper == new_fp.upper():
                        continue
                    if not Path(new_fp).exists():
                        continue

                    old_fp = orig_case.get(orig_upper, orig_upper)
                    try:
                        ok = sw.ReplaceReferencedDocument(asm_str, old_fp, new_fp)
                        if ok:
                            changed = True
                            logging.info(
                                "ReplaceRef OK: %s -> %s (%s)",
                                Path(old_fp).name, Path(new_fp).name, asm_ws.name
                            )
                    except Exception as e_rep:
                        logging.warning(
                            "ReplaceRef errore %s -> %s su %s: %s",
                            Path(old_fp).name, Path(new_fp).name, asm_ws.name, e_rep
                        )

                if changed:
                    modified += 1

            except Exception as e_asm:
                logging.warning("Fix riferimenti ASM %s: %s", asm_ws.name, e_asm)

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

    def _ensure_bom_document_map(self,
                                 path_to_code: dict[str, tuple[str, int, str]],
                                 created_doc_ids: list[int],
                                 created_doc_id_set: set[int]) -> tuple[int, int]:
        """
        Garantisce che ogni nodo dell'ASM abbia un documento PDM associato.
        - Se il nodo e gia mappato: nessuna azione.
        - Se esiste gia un documento con stesso code/doc_type: riusa.
        - Altrimenti crea un documento non codificato (code=stem originale, rev=00).
        Ritorna (riusati, creati).
        """
        reused = 0
        created = 0

        from config import SW_EXTENSIONS

        for node in self._nodes:
            p = node.get("path")
            if not p:
                continue

            key = str(p).upper()
            if key in path_to_code:
                continue

            fp = Path(p)
            code = fp.stem
            ext = fp.suffix if fp.suffix else ".SLDPRT"
            doc_type = SW_EXTENSIONS.get(ext.upper(), "")
            if not doc_type:
                doc_type = "Assieme" if node.get("type") == "ASM" else "Parte"

            existing = self.db.fetchone(
                "SELECT id, code, file_ext FROM documents "
                "WHERE code=? AND doc_type=? AND state != 'Obsoleto' "
                "ORDER BY revision DESC",
                (code, doc_type),
            )
            if existing:
                path_to_code[key] = (
                    existing["code"],
                    existing["id"],
                    existing.get("file_ext") or ext,
                )
                reused += 1
                continue

            # Crea documento non codificato per consentire il linking BOM completo.
            doc_id = self.files.create_document(
                code=code,
                revision="00",
                doc_type=doc_type,
                title=code,
                description="Creato automaticamente da import ASM (non codificato).",
                machine_id=None,
                group_id=None,
                doc_level=2,
            )
            created_doc_ids.append(doc_id)
            created_doc_id_set.add(doc_id)
            path_to_code[key] = (code, doc_id, ext)
            created += 1

        return reused, created


# =============================================================================
#  ui/coding_dialog.py  –  Gestione codifica gerarchica (Macchine / Gruppi)
# =============================================================================
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QGroupBox, QFormLayout, QTabWidget, QWidget,
    QTableWidget, QTableWidgetItem, QHeaderView, QComboBox,
    QMessageBox, QAbstractItemView
)
from PyQt6.QtCore import Qt

from ui.session import session


class CodingDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Macchine e Gruppi")
        self.setMinimumSize(700, 520)
        self._sel_machine_id = None
        self._sel_group_id   = None
        self._build_ui()
        self._load_machines()

    # ==================================================================
    # UI
    # ==================================================================
    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        lbl = QLabel(
            "Gestisci macchine, gruppi e contatori per la codifica gerarchica."
        )
        lbl.setObjectName("subtitle_label")
        layout.addWidget(lbl)

        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_machines_tab(), "Macchine")
        self.tabs.addTab(self._build_groups_tab(),   "Gruppi")
        self.tabs.addTab(self._build_counters_tab(), "Contatori")
        self.tabs.currentChanged.connect(self._tab_changed)
        layout.addWidget(self.tabs)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_close = QPushButton("Chiudi")
        btn_close.clicked.connect(self.accept)
        btn_row.addWidget(btn_close)
        layout.addLayout(btn_row)

    # ------------------------------------------------------------------
    # TAB – MACCHINE
    # ------------------------------------------------------------------
    def _build_machines_tab(self) -> QWidget:
        w = QWidget()
        layout = QHBoxLayout(w)

        # Lista macchine
        left = QVBoxLayout()
        self.tbl_machines = QTableWidget(0, 3)
        self.tbl_machines.setHorizontalHeaderLabels(["Codice", "Descrizione", "Tipo"])
        self.tbl_machines.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeMode.Stretch
        )
        self.tbl_machines.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self.tbl_machines.setEditTriggers(
            QAbstractItemView.EditTrigger.NoEditTriggers
        )
        self.tbl_machines.verticalHeader().setVisible(False)
        self.tbl_machines.selectionModel().selectionChanged.connect(
            self._on_machine_selected
        )
        left.addWidget(self.tbl_machines)

        layout.addLayout(left, stretch=2)

        # Pannello form
        right = QVBoxLayout()
        grp = QGroupBox("Macchina")
        form = QFormLayout(grp)

        self.m_code = QLineEdit()
        self.m_code.setMaximumWidth(120)
        self.m_code.textChanged.connect(
            lambda t: self.m_code.setText(t.upper()) if t != t.upper() else None
        )
        form.addRow("Codice:", self.m_code)

        self.m_desc = QLineEdit()
        self.m_desc.setPlaceholderText("Descrizione macchina")
        self.m_desc.textChanged.connect(
            lambda t: self.m_desc.setText(t.upper()) if t != t.upper() else None
        )
        form.addRow("Descrizione:", self.m_desc)

        right.addWidget(grp)

        r = QHBoxLayout()
        btn_new_m  = QPushButton("Nuova")
        btn_save_m = QPushButton("Salva")
        btn_save_m.setObjectName("btn_primary")
        btn_del_m  = QPushButton("Disattiva")
        btn_del_m.setObjectName("btn_warning")
        btn_new_m.clicked.connect(self._new_machine)
        btn_save_m.clicked.connect(self._save_machine)
        btn_del_m.clicked.connect(self._deactivate_machine)
        r.addWidget(btn_new_m)
        r.addStretch()
        r.addWidget(btn_del_m)
        r.addWidget(btn_save_m)
        right.addLayout(r)
        right.addStretch()

        layout.addLayout(right, stretch=1)
        return w

    # ------------------------------------------------------------------
    # TAB – GRUPPI
    # ------------------------------------------------------------------
    def _build_groups_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)

        # Selettore macchina
        top = QHBoxLayout()
        top.addWidget(QLabel("Macchina:"))
        self.cb_machine_grp = QComboBox()
        self.cb_machine_grp.setMinimumWidth(200)
        self.cb_machine_grp.currentIndexChanged.connect(self._on_machine_combo_changed)
        top.addWidget(self.cb_machine_grp)
        top.addStretch()
        layout.addLayout(top)

        split = QHBoxLayout()

        # Lista gruppi
        left = QVBoxLayout()
        self.tbl_groups = QTableWidget(0, 3)
        self.tbl_groups.setHorizontalHeaderLabels(["Codice", "Descrizione", "Tipo"])
        self.tbl_groups.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeMode.Stretch
        )
        self.tbl_groups.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self.tbl_groups.setEditTriggers(
            QAbstractItemView.EditTrigger.NoEditTriggers
        )
        self.tbl_groups.verticalHeader().setVisible(False)
        self.tbl_groups.selectionModel().selectionChanged.connect(
            self._on_group_selected
        )
        left.addWidget(self.tbl_groups)
        split.addLayout(left, stretch=2)

        # Form gruppo
        right = QVBoxLayout()
        grp = QGroupBox("Gruppo")
        form = QFormLayout(grp)

        self.g_code = QLineEdit()
        self.g_code.setMaximumWidth(120)
        self.g_code.textChanged.connect(
            lambda t: self.g_code.setText(t.upper()) if t != t.upper() else None
        )
        form.addRow("Codice:", self.g_code)

        self.g_desc = QLineEdit()
        self.g_desc.setPlaceholderText("Descrizione gruppo")
        self.g_desc.textChanged.connect(
            lambda t: self.g_desc.setText(t.upper()) if t != t.upper() else None
        )
        form.addRow("Descrizione:", self.g_desc)

        right.addWidget(grp)

        r = QHBoxLayout()
        btn_new_g  = QPushButton("Nuovo")
        btn_save_g = QPushButton("Salva")
        btn_save_g.setObjectName("btn_primary")
        btn_del_g  = QPushButton("Disattiva")
        btn_del_g.setObjectName("btn_warning")
        btn_new_g.clicked.connect(self._new_group)
        btn_save_g.clicked.connect(self._save_group)
        btn_del_g.clicked.connect(self._deactivate_group)
        r.addWidget(btn_new_g)
        r.addStretch()
        r.addWidget(btn_del_g)
        r.addWidget(btn_save_g)
        right.addLayout(r)
        right.addStretch()

        split.addLayout(right, stretch=1)
        layout.addLayout(split)
        return w

    # ------------------------------------------------------------------
    # TAB – CONTATORI
    # ------------------------------------------------------------------
    def _build_counters_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)

        btn_refresh = QPushButton("Aggiorna")
        btn_refresh.clicked.connect(self._load_counters)
        layout.addWidget(btn_refresh, alignment=Qt.AlignmentFlag.AlignRight)

        self.tbl_counters = QTableWidget(0, 5)
        self.tbl_counters.setHorizontalHeaderLabels(
            ["Tipo", "Macchina", "Gruppo", "Valore", ""]
        )
        self.tbl_counters.horizontalHeader().setSectionResizeMode(
            2, QHeaderView.ResizeMode.Stretch
        )
        self.tbl_counters.setEditTriggers(
            QAbstractItemView.EditTrigger.NoEditTriggers
        )
        self.tbl_counters.verticalHeader().setVisible(False)
        layout.addWidget(self.tbl_counters)

        note = QLabel(
            "⚠  Il reset manuale dei contatori può generare codici duplicati."
        )
        note.setObjectName("subtitle_label")
        layout.addWidget(note)
        return w

    # ==================================================================
    # LOGICA – MACCHINE
    # ==================================================================
    def _load_machines(self):
        machines = session.coding.get_machines(only_active=False)
        self.tbl_machines.setRowCount(0)
        self.tbl_machines.setProperty("_data", machines)
        for m in machines:
            r = self.tbl_machines.rowCount()
            self.tbl_machines.insertRow(r)
            self.tbl_machines.setItem(r, 0, QTableWidgetItem(m["code"]))
            self.tbl_machines.setItem(r, 1, QTableWidgetItem(m["description"] or ""))
            tipo = f"{m['code_type']} ({m['code_length']})"
            self.tbl_machines.setItem(r, 2, QTableWidgetItem(tipo))
            if not m["active"]:
                for c in range(3):
                    self.tbl_machines.item(r, c).setForeground(
                        Qt.GlobalColor.darkGray
                    )

        # Aggiorna anche il combobox gruppi
        self.cb_machine_grp.blockSignals(True)
        prev = self.cb_machine_grp.currentData()
        self.cb_machine_grp.clear()
        for m in machines:
            if m["active"]:
                self.cb_machine_grp.addItem(
                    f"{m['code']} – {m['description'] or ''}", m["id"]
                )
        if prev:
            idx = self.cb_machine_grp.findData(prev)
            if idx >= 0:
                self.cb_machine_grp.setCurrentIndex(idx)
        self.cb_machine_grp.blockSignals(False)
        self._on_machine_combo_changed()

    def _on_machine_selected(self):
        rows = self.tbl_machines.selectionModel().selectedRows()
        if not rows:
            self._sel_machine_id = None
            return
        row = rows[0].row()
        machines = session.coding.get_machines(only_active=False)
        if row < len(machines):
            m = machines[row]
            self._sel_machine_id = m["id"]
            self.m_code.setText(m["code"])
            self.m_desc.setText(m["description"] or "")

    def _new_machine(self):
        self.tbl_machines.clearSelection()
        self._sel_machine_id = None
        self.m_code.clear(); self.m_desc.clear()
        self.m_code.setFocus()

    def _save_machine(self):
        cfg   = session.coding.get_scheme_config()
        code  = self.m_code.text().strip().upper()
        desc  = self.m_desc.text().strip()
        ctype = cfg.mach_code_type
        clen  = cfg.mach_code_length

        ok, msg = session.coding.validate_code_string(code, ctype, clen)
        if not ok:
            QMessageBox.warning(self, "Codice non valido", msg)
            return

        try:
            if self._sel_machine_id:
                session.coding.update_machine(self._sel_machine_id, desc, ctype, clen)
            else:
                session.coding.create_machine(code, desc, ctype, clen)
            self._load_machines()
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

    def _deactivate_machine(self):
        if not self._sel_machine_id:
            return
        r = QMessageBox.question(
            self, "Disattiva macchina",
            "La macchina non apparirà più nella selezione. Continuare?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if r == QMessageBox.StandardButton.Yes:
            session.coding.deactivate_machine(self._sel_machine_id)
            self._sel_machine_id = None
            self._load_machines()

    # ==================================================================
    # LOGICA – GRUPPI
    # ==================================================================
    def _on_machine_combo_changed(self):
        machine_id = self.cb_machine_grp.currentData()
        self._load_groups(machine_id)

    def _load_groups(self, machine_id):
        self.tbl_groups.setRowCount(0)
        if not machine_id:
            return
        groups = session.coding.get_groups(machine_id, only_active=False)
        self.tbl_groups.setProperty("_data", groups)
        for g in groups:
            r = self.tbl_groups.rowCount()
            self.tbl_groups.insertRow(r)
            self.tbl_groups.setItem(r, 0, QTableWidgetItem(g["code"]))
            self.tbl_groups.setItem(r, 1, QTableWidgetItem(g["description"] or ""))
            tipo = f"{g['code_type']} ({g['code_length']})"
            self.tbl_groups.setItem(r, 2, QTableWidgetItem(tipo))
            if not g["active"]:
                for c in range(3):
                    self.tbl_groups.item(r, c).setForeground(
                        Qt.GlobalColor.darkGray
                    )

    def _on_group_selected(self):
        rows = self.tbl_groups.selectionModel().selectedRows()
        if not rows:
            self._sel_group_id = None
            return
        row = rows[0].row()
        machine_id = self.cb_machine_grp.currentData()
        if not machine_id:
            return
        groups = session.coding.get_groups(machine_id, only_active=False)
        if row < len(groups):
            g = groups[row]
            self._sel_group_id = g["id"]
            self.g_code.setText(g["code"])
            self.g_desc.setText(g["description"] or "")

    def _new_group(self):
        self.tbl_groups.clearSelection()
        self._sel_group_id = None
        self.g_code.clear(); self.g_desc.clear()
        self.g_code.setFocus()

    def _save_group(self):
        machine_id = self.cb_machine_grp.currentData()
        if not machine_id:
            QMessageBox.warning(self, "Attenzione", "Seleziona prima una macchina")
            return

        cfg   = session.coding.get_scheme_config()
        code  = self.g_code.text().strip().upper()
        desc  = self.g_desc.text().strip()
        ctype = cfg.grp_code_type
        clen  = cfg.grp_code_length

        ok, msg = session.coding.validate_code_string(code, ctype, clen)
        if not ok:
            QMessageBox.warning(self, "Codice non valido", msg)
            return

        try:
            if self._sel_group_id:
                session.coding.update_group(self._sel_group_id, desc, ctype, clen)
            else:
                session.coding.create_group(machine_id, code, desc, ctype, clen)
            self._load_groups(machine_id)
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

    def _deactivate_group(self):
        if not self._sel_group_id:
            return
        r = QMessageBox.question(
            self, "Disattiva gruppo",
            "Il gruppo non apparirà più nella selezione. Continuare?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if r == QMessageBox.StandardButton.Yes:
            session.coding.deactivate_group(self._sel_group_id)
            self._sel_group_id = None
            machine_id = self.cb_machine_grp.currentData()
            self._load_groups(machine_id)

    # ==================================================================
    # LOGICA – CONTATORI
    # ==================================================================
    def _load_counters(self):
        counters = session.coding.get_counters()
        self.tbl_counters.setRowCount(0)
        is_admin = (session.user or {}).get("role", "") == "Admin"
        for ct in counters:
            r = self.tbl_counters.rowCount()
            self.tbl_counters.insertRow(r)
            self.tbl_counters.setItem(r, 0, QTableWidgetItem(ct["counter_type"]))
            self.tbl_counters.setItem(r, 1, QTableWidgetItem(ct.get("machine_code") or ""))
            self.tbl_counters.setItem(r, 2, QTableWidgetItem(ct.get("group_code") or ""))
            self.tbl_counters.setItem(r, 3, QTableWidgetItem(str(ct["last_value"])))
            if is_admin:
                btn = QPushButton("Reset")
                btn.setObjectName("btn_warning")
                btn.clicked.connect(
                    lambda _, row=ct: self._reset_counter_row(row)
                )
                self.tbl_counters.setCellWidget(r, 4, btn)

    def _reset_counter_row(self, row: dict):
        r = QMessageBox.question(
            self, "Reset contatore",
            f"Resettare il contatore {row['counter_type']} "
            f"({row.get('machine_code','')} / {row.get('group_code','')}) a 0?\n"
            "Attenzione: potrebbero generarsi codici duplicati!",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if r == QMessageBox.StandardButton.Yes:
            session.coding.reset_counter(
                row["counter_type"], row["machine_id"], row["group_id"], 0
            )
            self._load_counters()

    # Compatibilità: ricarica quando si passa alla tab contatori
    def _tab_changed(self, idx: int):
        if idx == 2:
            self._load_counters()


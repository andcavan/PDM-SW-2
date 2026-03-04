# =============================================================================
#  ui/users_dialog.py  –  Gestione utenti (solo Amministratore)
# =============================================================================
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
    QComboBox, QGroupBox, QFormLayout, QMessageBox, QCheckBox
)
from PyQt6.QtCore import Qt

from config import USER_ROLES
from ui.session import session


class UsersDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Gestione Utenti")
        self.setMinimumSize(700, 480)
        self._editing_id = None
        self._build_ui()
        self._refresh()

    def _build_ui(self):
        layout = QHBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(16, 16, 16, 16)

        # Lista utenti
        left = QVBoxLayout()
        left.addWidget(QLabel("Utenti registrati:"))
        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["Utente", "Nome", "Ruolo"])
        self.table.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeMode.Stretch
        )
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.selectionModel().selectionChanged.connect(self._on_select)
        left.addWidget(self.table)
        btn_new = QPushButton("+ Nuovo utente")
        btn_new.clicked.connect(self._new_user)
        btn_del = QPushButton("Disattiva utente")
        btn_del.setObjectName("btn_danger")
        btn_del.clicked.connect(self._delete_user)
        row = QHBoxLayout()
        row.addWidget(btn_new)
        row.addWidget(btn_del)
        left.addLayout(row)
        layout.addLayout(left, 3)

        # Form dettaglio
        right = QVBoxLayout()
        grp = QGroupBox("Dettaglio utente")
        form = QFormLayout(grp)

        self.txt_username  = QLineEdit()
        self.txt_fullname  = QLineEdit()
        self.cmb_role      = QComboBox()
        self.cmb_role.addItems(USER_ROLES)
        self.txt_password  = QLineEdit()
        self.txt_password.setEchoMode(QLineEdit.EchoMode.Password)
        self.txt_password.setPlaceholderText("Lascia vuoto per non cambiare")
        self.chk_active    = QCheckBox("Attivo")
        self.chk_active.setChecked(True)

        form.addRow("Username:", self.txt_username)
        form.addRow("Nome completo:", self.txt_fullname)
        form.addRow("Ruolo:", self.cmb_role)
        form.addRow("Password:", self.txt_password)
        form.addRow(self.chk_active)

        right.addWidget(grp)
        right.addStretch()

        btn_save = QPushButton("Salva")
        btn_save.setObjectName("btn_primary")
        btn_save.clicked.connect(self._save_user)
        right.addWidget(btn_save)

        layout.addLayout(right, 2)

    # ------------------------------------------------------------------
    def _refresh(self):
        users = session.users.get_all_users()
        self.table.setRowCount(0)
        for u in users:
            r = self.table.rowCount()
            self.table.insertRow(r)
            self.table.setItem(r, 0, QTableWidgetItem(u["username"]))
            self.table.setItem(r, 1, QTableWidgetItem(u["full_name"]))
            self.table.setItem(r, 2, QTableWidgetItem(u["role"]))
            self.table.item(r, 0).setData(Qt.ItemDataRole.UserRole, u["id"])

    def _on_select(self):
        row = self.table.currentRow()
        if row < 0:
            return
        uid = self.table.item(row, 0).data(Qt.ItemDataRole.UserRole)
        u   = session.users.get_user(uid)
        if not u:
            return
        self._editing_id = uid
        self.txt_username.setText(u["username"])
        self.txt_username.setReadOnly(u["username"] == "admin")
        self.txt_fullname.setText(u["full_name"])
        idx = self.cmb_role.findText(u["role"])
        if idx >= 0:
            self.cmb_role.setCurrentIndex(idx)
        self.chk_active.setChecked(bool(u["active"]))
        self.txt_password.clear()

    def _new_user(self):
        self._editing_id = None
        self.txt_username.clear()
        self.txt_username.setReadOnly(False)
        self.txt_fullname.clear()
        self.cmb_role.setCurrentText("Progettista")
        self.chk_active.setChecked(True)
        self.txt_password.clear()
        self.txt_username.setFocus()

    def _save_user(self):
        username  = self.txt_username.text().strip()
        full_name = self.txt_fullname.text().strip()
        role      = self.cmb_role.currentText()
        password  = self.txt_password.text()
        active    = self.chk_active.isChecked()

        if not username or not full_name:
            QMessageBox.warning(self, "Errore", "Username e Nome completo sono obbligatori")
            return

        try:
            if self._editing_id:
                session.users.update_user(
                    self._editing_id, full_name, role, password, active
                )
            else:
                session.users.create_user(username, full_name, role, password)
            self._refresh()
            QMessageBox.information(self, "OK", "Utente salvato")
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

    def _delete_user(self):
        row = self.table.currentRow()
        if row < 0:
            return
        uid  = self.table.item(row, 0).data(Qt.ItemDataRole.UserRole)
        name = self.table.item(row, 1).text()
        r = QMessageBox.question(
            self, "Disattiva utente",
            f"Disattivare l'utente '{name}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if r == QMessageBox.StandardButton.Yes:
            session.users.delete_user(uid)
            self._refresh()

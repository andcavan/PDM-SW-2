# =============================================================================
#  ui/login_dialog.py  –  Selezione utente e login
# =============================================================================
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QComboBox, QMessageBox, QFrame
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont

from config import APP_NAME, APP_VERSION
from ui.session import session


class LoginDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"{APP_NAME} – Accesso")
        self.setFixedWidth(360)
        self.setModal(True)
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        self._build_ui()
        self._load_users()

    # ------------------------------------------------------------------
    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(14)
        layout.setContentsMargins(30, 30, 30, 30)

        # Logo / titolo
        lbl_title = QLabel(f"🔧  {APP_NAME}")
        lbl_title.setObjectName("title_label")
        lbl_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(lbl_title)

        lbl_ver = QLabel(f"v{APP_VERSION}")
        lbl_ver.setObjectName("subtitle_label")
        lbl_ver.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(lbl_ver)

        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color:#313244"); layout.addWidget(sep)

        # Utente
        layout.addWidget(QLabel("Utente:"))
        self.cmb_user = QComboBox()
        self.cmb_user.setEditable(False)
        self.cmb_user.currentIndexChanged.connect(self._on_user_changed)
        layout.addWidget(self.cmb_user)

        # Password
        self.lbl_pwd = QLabel("Password:")
        self.txt_pwd = QLineEdit()
        self.txt_pwd.setEchoMode(QLineEdit.EchoMode.Password)
        self.txt_pwd.setPlaceholderText("Lascia vuoto se non richiesta")
        self.txt_pwd.returnPressed.connect(self._login)
        layout.addWidget(self.lbl_pwd)
        layout.addWidget(self.txt_pwd)

        # Bottoni
        btn_row = QHBoxLayout()
        btn_login = QPushButton("Accedi")
        btn_login.setObjectName("btn_primary")
        btn_login.clicked.connect(self._login)
        btn_cancel = QPushButton("Annulla")
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_login)
        layout.addLayout(btn_row)

        # No-password hint
        cfg = __import__("config").load_local_config()
        if cfg.get("no_password", True):
            hint = QLabel("ℹ️  Modalità rete: password non richiesta")
            hint.setObjectName("subtitle_label")
            hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(hint)
            self.lbl_pwd.hide()
            self.txt_pwd.hide()
            self._no_pwd = True
        else:
            self._no_pwd = False

    # ------------------------------------------------------------------
    def _load_users(self):
        users = session.users.get_all_users()
        for u in users:
            self.cmb_user.addItem(
                f"{u['full_name']} ({u['username']})", u["username"]
            )

    def _on_user_changed(self, idx):
        pass

    # ------------------------------------------------------------------
    def _login(self):
        username = self.cmb_user.currentData()
        if not username:
            QMessageBox.warning(self, "Errore", "Selezionare un utente")
            return

        if self._no_pwd:
            user = session.users.login_no_password(username)
        else:
            pwd  = self.txt_pwd.text()
            user = session.users.authenticate(username, pwd)

        if not user:
            QMessageBox.warning(
                self, "Accesso negato",
                "Credenziali non valide o utente non attivo."
            )
            return

        session.set_user(user)
        self.accept()

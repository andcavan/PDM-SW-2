# =============================================================================
#  ui/setup_dialog.py  –  Setup iniziale / configurazione percorso condiviso
# =============================================================================
from pathlib import Path
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QFileDialog, QMessageBox, QGroupBox, QFormLayout,
    QCheckBox, QFrame
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont

from config import APP_NAME, APP_VERSION, save_local_config, load_local_config


class SetupDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"{APP_NAME} – Configurazione Iniziale")
        self.setMinimumWidth(520)
        self.setModal(True)
        self._build_ui()
        self._load_existing()

    # ------------------------------------------------------------------
    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(16)
        layout.setContentsMargins(24, 24, 24, 24)

        # Titolo
        title = QLabel(f"⚙️  {APP_NAME}  v{APP_VERSION}")
        title.setObjectName("title_label")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        sub = QLabel("Configurazione percorso cartella condivisa di rete")
        sub.setObjectName("subtitle_label")
        sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(sub)

        # Separatore
        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color: #313244;"); layout.addWidget(sep)

        # Percorso condiviso
        grp = QGroupBox("Cartella condivisa (UNC o mappata)")
        form = QFormLayout(grp)

        row = QHBoxLayout()
        self.txt_shared = QLineEdit()
        self.txt_shared.setPlaceholderText(r"es. \\SERVER\PDM  oppure  Z:\PDM")
        btn_browse = QPushButton("Sfoglia…")
        btn_browse.clicked.connect(self._browse_shared)
        row.addWidget(self.txt_shared)
        row.addWidget(btn_browse)
        form.addRow("Percorso:", row)
        layout.addWidget(grp)

        # Impostazioni locali
        grp2 = QGroupBox("Impostazioni locali")
        form2 = QFormLayout(grp2)
        self.chk_no_pwd = QCheckBox(
            "Login senza password (ambiente protetto di rete)"
        )
        form2.addRow(self.chk_no_pwd)
        layout.addWidget(grp2)

        # Bottoni
        btn_row = QHBoxLayout()
        self.btn_test = QPushButton("Verifica connessione")
        self.btn_test.clicked.connect(self._test_connection)
        btn_row.addWidget(self.btn_test)
        btn_row.addStretch()
        btn_save = QPushButton("Salva e Connetti")
        btn_save.setObjectName("btn_primary")
        btn_save.clicked.connect(self._save)
        btn_cancel = QPushButton("Annulla")
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_save)
        layout.addLayout(btn_row)

    # ------------------------------------------------------------------
    def _load_existing(self):
        cfg = load_local_config()
        self.txt_shared.setText(cfg.get("shared_root", ""))
        self.chk_no_pwd.setChecked(cfg.get("no_password", True))

    def _browse_shared(self):
        path = QFileDialog.getExistingDirectory(
            self, "Seleziona cartella condivisa PDM"
        )
        if path:
            self.txt_shared.setText(path)

    def _test_connection(self):
        path = Path(self.txt_shared.text().strip())
        if path.exists():
            QMessageBox.information(
                self, "Connessione OK",
                f"Cartella raggiungibile:\n{path}"
            )
        else:
            QMessageBox.warning(
                self, "Errore",
                f"Cartella non raggiungibile:\n{path}"
            )

    def _save(self):
        shared = self.txt_shared.text().strip()
        if not shared:
            QMessageBox.warning(self, "Errore", "Inserire il percorso della cartella condivisa")
            return
        path = Path(shared)
        if not path.exists():
            r = QMessageBox.question(
                self, "Cartella non trovata",
                "La cartella non esiste. Crearla?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if r == QMessageBox.StandardButton.Yes:
                path.mkdir(parents=True, exist_ok=True)
            else:
                return

        save_local_config({
            "shared_root": str(path),
            "no_password": self.chk_no_pwd.isChecked(),
        })
        self.accept()

    @property
    def shared_root(self) -> str:
        return self.txt_shared.text().strip()

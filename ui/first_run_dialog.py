# =============================================================================
#  ui/first_run_dialog.py  –  Dialog primo avvio: cartella dati + workspace SW
# =============================================================================
import json
from pathlib import Path

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QFileDialog, QMessageBox, QGroupBox, QFormLayout, QFrame,
)
from PyQt6.QtCore import Qt

from config import APP_NAME, APP_VERSION, DATA_DIR_FILE


class FirstRunDialog(QDialog):
    """Mostrato al primo avvio per scegliere la cartella dati locali e il workspace SW."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"{APP_NAME} – Primo avvio")
        self.setMinimumWidth(580)
        self.setModal(True)
        self._build_ui()

    # ------------------------------------------------------------------
    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(16)
        layout.setContentsMargins(28, 28, 28, 28)

        title = QLabel(f"⚙️  Benvenuto in {APP_NAME}  v{APP_VERSION}")
        title.setObjectName("title_label")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        sub = QLabel(
            "Prima configurazione – scegli dove salvare i dati locali "
            "e la cartella workspace SolidWorks."
        )
        sub.setObjectName("subtitle_label")
        sub.setWordWrap(True)
        sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(sub)

        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color: #313244;")
        layout.addWidget(sep)

        # -- Cartella dati locali --
        grp1 = QGroupBox("Cartella dati locali")
        form1 = QFormLayout(grp1)

        row1 = QHBoxLayout()
        self.txt_datadir = QLineEdit()
        self.txt_datadir.setPlaceholderText(r"es. C:\Users\mario\PDM-Data")
        btn1 = QPushButton("Sfoglia…")
        btn1.clicked.connect(
            lambda: self._browse(self.txt_datadir, "Seleziona cartella dati locali")
        )
        row1.addWidget(self.txt_datadir)
        row1.addWidget(btn1)
        form1.addRow("Percorso:", row1)

        note1 = QLabel(
            "Contiene local_config.json con le impostazioni di questo PC. "
            "Non deve essere condivisa in rete."
        )
        note1.setObjectName("subtitle_label")
        note1.setWordWrap(True)
        form1.addRow(note1)
        layout.addWidget(grp1)

        # -- Workspace SolidWorks --
        grp2 = QGroupBox("Workspace SolidWorks (checkout locale)")
        form2 = QFormLayout(grp2)

        row2 = QHBoxLayout()
        self.txt_workspace = QLineEdit()
        self.txt_workspace.setPlaceholderText(r"es. C:\Lavoro\PDM_WS")
        btn2 = QPushButton("Sfoglia…")
        btn2.clicked.connect(
            lambda: self._browse(self.txt_workspace, "Seleziona cartella workspace SolidWorks")
        )
        row2.addWidget(self.txt_workspace)
        row2.addWidget(btn2)
        form2.addRow("Percorso:", row2)

        note2 = QLabel(
            "Cartella locale dove vengono copiati i file in checkout. "
            "Deve essere accessibile in lettura/scrittura."
        )
        note2.setObjectName("subtitle_label")
        note2.setWordWrap(True)
        form2.addRow(note2)
        layout.addWidget(grp2)

        # -- Bottoni --
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_cancel = QPushButton("Annulla")
        btn_cancel.clicked.connect(self.reject)
        btn_ok = QPushButton("Conferma e continua →")
        btn_ok.setObjectName("btn_primary")
        btn_ok.clicked.connect(self._accept)
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_ok)
        layout.addLayout(btn_row)

    # ------------------------------------------------------------------
    def _browse(self, target: QLineEdit, title: str):
        path = QFileDialog.getExistingDirectory(self, title)
        if path:
            target.setText(path)

    def _accept(self):
        data_dir = self.txt_datadir.text().strip()
        workspace = self.txt_workspace.text().strip()

        if not data_dir:
            QMessageBox.warning(self, "Campo mancante", "Selezionare la cartella dati locali.")
            return
        if not workspace:
            QMessageBox.warning(self, "Campo mancante", "Selezionare la cartella workspace SolidWorks.")
            return

        data_path = Path(data_dir)
        ws_path = Path(workspace)

        try:
            data_path.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            QMessageBox.critical(self, "Errore", f"Impossibile creare la cartella dati:\n{e}")
            return

        try:
            ws_path.mkdir(parents=True, exist_ok=True)
        except OSError:
            pass  # Non blocca: può essere su rete non ancora disponibile

        # Scrivi .pdm_datadir (puntatore alla cartella dati)
        DATA_DIR_FILE.write_text(str(data_path), encoding="utf-8")

        # Inizializza local_config.json solo se non esiste già
        cfg_file = data_path / "local_config.json"
        if not cfg_file.exists():
            init_data = {
                "active_profile": "",
                "profiles": {},
                "_init_workspace": str(ws_path),
            }
            cfg_file.write_text(
                json.dumps(init_data, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )

        self.accept()

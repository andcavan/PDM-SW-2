# =============================================================================
#  ui/commercial_settings_dialog.py  –  Impostazioni modulo commerciali
# =============================================================================
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QDialogButtonBox, QFileDialog, QGroupBox,
    QFormLayout,
)
from PyQt6.QtCore import Qt

from ui.session import session


class CommercialSettingsDialog(QDialog):
    """Dialog per configurare le impostazioni del modulo articoli commerciali."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Impostazioni commerciali")
        self.setMinimumWidth(520)
        self._build_ui()
        self._load()

    # ------------------------------------------------------------------

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setSpacing(12)
        lay.setContentsMargins(14, 14, 14, 14)

        # Gruppo percorso archivio
        grp = QGroupBox("Archivio file SolidWorks")
        form = QFormLayout(grp)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        form.setSpacing(8)

        path_row = QHBoxLayout()
        self.txt_archive_path = QLineEdit()
        self.txt_archive_path.setPlaceholderText(
            "Lasciare vuoto per usare la cartella predefinita"
        )
        path_row.addWidget(self.txt_archive_path)
        btn_browse = QPushButton("Sfoglia…")
        btn_browse.setFixedWidth(80)
        btn_browse.clicked.connect(self._browse)
        path_row.addWidget(btn_browse)
        form.addRow("Percorso archivio:", path_row)

        lbl_hint = QLabel(
            "Cartella in cui vengono archiviati i file SW degli articoli commerciali.\n"
            "Se vuoto: usa la sottocartella predefinita dell'archivio condiviso.\n"
            "Può essere un percorso assoluto (es. Z:\\commerciali) o relativo alla radice condivisa."
        )
        lbl_hint.setWordWrap(True)
        lbl_hint.setStyleSheet("color: #888; font-size: 11px;")
        form.addRow("", lbl_hint)

        lay.addWidget(grp)

        # Bottoni
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save
            | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.button(QDialogButtonBox.StandardButton.Save).setText("Salva")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Annulla")
        buttons.accepted.connect(self._save)
        buttons.rejected.connect(self.reject)
        lay.addWidget(buttons)

    # ------------------------------------------------------------------

    def _browse(self):
        folder = QFileDialog.getExistingDirectory(
            self, "Seleziona cartella archivio commerciali",
            self.txt_archive_path.text() or "",
        )
        if folder:
            self.txt_archive_path.setText(folder)

    def _load(self):
        if not session.commercial:
            return
        cfg = session.commercial.get_coding_config()
        self.txt_archive_path.setText(cfg.commercial_archive_path or "")

    def _save(self):
        if not session.commercial:
            self.accept()
            return
        cfg = session.commercial.get_coding_config()
        cfg.commercial_archive_path = self.txt_archive_path.text().strip()
        try:
            session.commercial.save_coding_config(cfg)
        except Exception as e:
            from PyQt6.QtWidgets import QMessageBox
            QMessageBox.critical(self, "Errore salvataggio", str(e))
            return
        self.accept()

#!/usr/bin/env python
# =============================================================================
#  setup_app.py  –  Setup/installazione iniziale PDM-SW
#  Eseguire UNA SOLA VOLTA sul PC che ospita la cartella condivisa
#  oppure dalla stazione di lavoro per la configurazione locale.
# =============================================================================
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PyQt6.QtWidgets import (
    QApplication, QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QLineEdit, QFileDialog, QComboBox, QGroupBox,
    QFormLayout, QMessageBox, QCheckBox, QTextEdit, QWizard,
    QWizardPage, QFrame
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont

from config import APP_NAME, APP_VERSION, SharedPaths, save_local_config
from ui.styles import DARK_THEME


class WelcomePage(QWizardPage):
    def __init__(self):
        super().__init__()
        self.setTitle(f"Benvenuto in {APP_NAME}")
        self.setSubTitle(f"Versione {APP_VERSION} – Wizard di configurazione iniziale")
        layout = QVBoxLayout(self)
        info = QLabel(
            f"<b>{APP_NAME}</b> è un sistema PDM per la gestione di documenti SolidWorks.<br><br>"
            "Questo wizard ti guiderà nella configurazione iniziale:<br>"
            "<ul>"
            "<li>Percorso della cartella condivisa di rete</li>"
            "<li>Creazione struttura archivio PDM</li>"
            "<li>Configurazione utente amministratore</li>"
            "<li>Impostazioni codifica documenti</li>"
            "</ul><br>"
            "<b>Requisiti:</b><br>"
            "• Python 3.10+ con pacchetti da requirements.txt installati<br>"
            "• Cartella di rete condivisa accessibile da tutti gli utenti (max 5)<br>"
            "• Permessi di lettura/scrittura sulla cartella condivisa<br>"
            "• SolidWorks 2018+ (opzionale, per integrazione COM)"
        )
        info.setWordWrap(True)
        info.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(info)


class SharedPathPage(QWizardPage):
    def __init__(self):
        super().__init__()
        self.setTitle("Percorso cartella condivisa")
        self.setSubTitle(
            "Inserisci il percorso UNC o mappato della cartella condivisa "
            "che ospiterà il database e l'archivio file PDM."
        )
        layout = QVBoxLayout(self)

        grp = QGroupBox("Cartella condivisa PDM")
        form = QFormLayout(grp)

        row = QHBoxLayout()
        self.txt_path = QLineEdit()
        self.txt_path.setPlaceholderText(r"es. \\SERVER\Condivisa\PDM  oppure  Z:\PDM")
        btn = QPushButton("Sfoglia…")
        btn.clicked.connect(self._browse)
        row.addWidget(self.txt_path)
        row.addWidget(btn)
        form.addRow("Percorso:", row)

        self.lbl_status = QLabel("")
        form.addRow("Stato:", self.lbl_status)

        btn_test = QPushButton("Verifica accessibilità")
        btn_test.clicked.connect(self._test)
        form.addRow(btn_test)

        layout.addWidget(grp)

        note = QLabel(
            "ℹ️  La struttura delle cartelle verrà creata automaticamente:\n"
            "  /database/  – Database SQLite PDM\n"
            "  /archive/   – Archivio file SolidWorks\n"
            "  /workspace/ – Workspace utenti\n"
            "  /thumbnails/– Anteprime documenti\n"
            "  /config/    – Configurazioni condivise\n"
            "  /temp/      – File temporanei"
        )
        note.setObjectName("subtitle_label")
        layout.addWidget(note)

        self.registerField("shared_path*", self.txt_path)

    def _browse(self):
        path = QFileDialog.getExistingDirectory(
            self, "Seleziona cartella condivisa"
        )
        if path:
            self.txt_path.setText(path)

    def _test(self):
        from pathlib import Path
        p = Path(self.txt_path.text().strip())
        if p.exists():
            self.lbl_status.setText("✅ Accessibile")
            self.lbl_status.setStyleSheet("color:#a6e3a1;")
        else:
            try:
                p.mkdir(parents=True, exist_ok=True)
                self.lbl_status.setText("✅ Creata e accessibile")
                self.lbl_status.setStyleSheet("color:#a6e3a1;")
            except Exception as e:
                self.lbl_status.setText(f"❌ Errore: {e}")
                self.lbl_status.setStyleSheet("color:#f38ba8;")


class AdminUserPage(QWizardPage):
    def __init__(self):
        super().__init__()
        self.setTitle("Utente Amministratore")
        self.setSubTitle(
            "Configura le credenziali dell'utente amministratore PDM. "
            "L'utente 'admin' è già creato di default (password: admin)."
        )
        layout = QVBoxLayout(self)

        grp = QGroupBox("Primo utente – Amministratore")
        form = QFormLayout(grp)

        self.txt_admin_name = QLineEdit("Amministratore PDM")
        form.addRow("Nome completo:", self.txt_admin_name)

        self.txt_admin_pwd = QLineEdit()
        self.txt_admin_pwd.setEchoMode(QLineEdit.EchoMode.Password)
        self.txt_admin_pwd.setPlaceholderText("Lascia vuoto per usare 'admin'")
        form.addRow("Password admin:", self.txt_admin_pwd)

        self.chk_no_pwd = QCheckBox(
            "Modalità rete sicura (login senza password)"
        )
        self.chk_no_pwd.setChecked(True)
        form.addRow(self.chk_no_pwd)

        layout.addWidget(grp)

        note = QLabel(
            "Potrai aggiungere altri utenti (max 5) dopo il setup.\n"
            "Ruoli disponibili: Utente, Progettista, Responsabile, Amministratore"
        )
        note.setObjectName("subtitle_label")
        layout.addWidget(note)

        self.registerField("admin_name", self.txt_admin_name)
        self.registerField("admin_pwd", self.txt_admin_pwd)


class CodingSetupPage(QWizardPage):
    def __init__(self):
        super().__init__()
        self.setTitle("Configurazione codifica documenti")
        self.setSubTitle(
            "Definisci il formato dei codici. Potrai modificarli in seguito "
            "da Strumenti → Configurazione codifica."
        )
        layout = QVBoxLayout(self)

        grp = QGroupBox("Formato codici predefiniti")
        form = QFormLayout(grp)

        self.cmb_scheme = QComboBox()
        self.cmb_scheme.addItems([
            "PP-00001 / AS-00001 / DP-00001  (predefinito)",
            "P-001 / A-001 / D-001  (compatto)",
            "PART-00001 / ASM-00001 / DRW-00001  (esteso)",
            "Personalizzato",
        ])
        self.cmb_scheme.currentIndexChanged.connect(self._update_preview)
        form.addRow("Schema:", self.cmb_scheme)

        self.lbl_preview_p = QLabel("PP-00001")
        self.lbl_preview_a = QLabel("AS-00001")
        self.lbl_preview_d = QLabel("DP-00001")
        for lbl in [self.lbl_preview_p, self.lbl_preview_a, self.lbl_preview_d]:
            lbl.setStyleSheet("color:#89b4fa;font-weight:bold;font-size:14px;")
        form.addRow("Anteprima Parte:", self.lbl_preview_p)
        form.addRow("Anteprima Assieme:", self.lbl_preview_a)
        form.addRow("Anteprima Disegno:", self.lbl_preview_d)

        layout.addWidget(grp)
        self.registerField("coding_scheme", self.cmb_scheme, "currentIndex")

    def _update_preview(self, idx):
        schemes = [
            ("PP", "AS", "DP", "-", 5),
            ("P",  "A",  "D",  "-", 3),
            ("PART", "ASM", "DRW", "-", 5),
            ("XX", "XX", "XX", "-", 5),
        ]
        p, a, d, sep, dig = schemes[min(idx, 3)]
        num = "1".zfill(dig)
        self.lbl_preview_p.setText(f"{p}{sep}{num}" if p else num)
        self.lbl_preview_a.setText(f"{a}{sep}{num}" if a else num)
        self.lbl_preview_d.setText(f"{d}{sep}{num}" if d else num)


class SummaryPage(QWizardPage):
    def __init__(self):
        super().__init__()
        self.setTitle("Riepilogo configurazione")
        self.setSubTitle("Controlla le impostazioni prima di applicare.")
        layout = QVBoxLayout(self)
        self.txt_summary = QTextEdit()
        self.txt_summary.setReadOnly(True)
        layout.addWidget(self.txt_summary)

    def initializePage(self):
        shared = self.field("shared_path") or "—"
        admin  = self.field("admin_name") or "Amministratore PDM"
        self.txt_summary.setPlainText(
            f"Cartella condivisa:\n  {shared}\n\n"
            f"Utente amministratore:\n  {admin}\n\n"
            "Struttura cartelle che verrà creata:\n"
            f"  {shared}\\database\\\n"
            f"  {shared}\\archive\\\n"
            f"  {shared}\\workspace\\\n"
            f"  {shared}\\thumbnails\\\n"
            f"  {shared}\\config\\\n"
            f"  {shared}\\temp\\\n\n"
            "Clicca 'Fine' per applicare la configurazione."
        )


class SetupWizard(QWizard):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} – Setup Iniziale")
        self.setMinimumSize(620, 480)
        self.setWizardStyle(QWizard.WizardStyle.ModernStyle)

        self.addPage(WelcomePage())
        self.addPage(SharedPathPage())
        self.addPage(AdminUserPage())
        self.addPage(CodingSetupPage())
        self.addPage(SummaryPage())

        self.button(QWizard.WizardButton.FinishButton).clicked.connect(
            self._apply_setup
        )

    def _apply_setup(self):
        from pathlib import Path
        from config import SharedPaths
        from core.database import Database
        from core.user_manager import UserManager
        from core.coding_manager import CodingManager

        shared_root = self.field("shared_path")
        admin_name  = self.field("admin_name") or "Amministratore PDM"
        admin_pwd   = self.field("admin_pwd") or "admin"

        try:
            sp = SharedPaths(shared_root)
            sp.ensure_dirs()

            db = Database(sp.db_file, sp.db_lock_file)
            db.initialize()

            # Aggiorna admin
            um = UserManager(db)
            admin = db.fetchone(
                "SELECT id FROM users WHERE username='admin'"
            )
            if admin:
                um.update_user(admin["id"], admin_name, "Amministratore", admin_pwd)

            # Configura codifica in base allo schema scelto
            scheme_idx = self.field("coding_scheme")
            cm = CodingManager(db)
            schemes = [
                [("Parte", "PP", "-", 5), ("Assieme", "AS", "-", 5), ("Disegno", "DP", "-", 5)],
                [("Parte", "P",  "-", 3), ("Assieme", "A",  "-", 3), ("Disegno", "D",  "-", 3)],
                [("Parte", "PART", "-", 5), ("Assieme", "ASM", "-", 5), ("Disegno", "DRW", "-", 5)],
                [("Parte", "PP", "-", 5), ("Assieme", "AS", "-", 5), ("Disegno", "DP", "-", 5)],
            ]
            for doc_type, prefix, sep, digits in schemes[min(scheme_idx, 3)]:
                cm.save_config(doc_type, prefix, sep, digits, 1, "")

            # Salva config locale
            from config import load_local_config
            cfg = load_local_config()
            cfg["shared_root"] = str(shared_root)
            save_local_config(cfg)

            QMessageBox.information(
                self, "Setup completato!",
                f"{APP_NAME} è stato configurato con successo.\n\n"
                f"Database creato in:\n{sp.db_file}\n\n"
                "Avviare main.py per iniziare."
            )

        except Exception as e:
            QMessageBox.critical(
                self, "Errore setup", f"Errore durante la configurazione:\n{e}"
            )


def main():
    app = QApplication(sys.argv)
    app.setApplicationName(f"{APP_NAME} Setup")
    app.setStyle("Fusion")
    app.setStyleSheet(DARK_THEME)
    app.setFont(QFont("Segoe UI", 10))

    wizard = SetupWizard()
    wizard.exec()
    sys.exit(0)


if __name__ == "__main__":
    main()

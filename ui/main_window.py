# =============================================================================
#  ui/main_window.py  –  Finestra principale PDM-SW
# =============================================================================
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QStatusBar, QLabel, QPushButton,
    QMessageBox, QToolBar, QSplitter
)
from PyQt6.QtCore import Qt, QTimer, QSettings, QByteArray
from PyQt6.QtGui import QAction, QFont, QKeySequence

from config import APP_NAME, APP_VERSION
from ui.session import session
from ui.archive_view import ArchiveView
from ui.workspace_view import WorkspaceView
from ui.document_dialog import DocumentDialog


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME}  v{APP_VERSION}")
        self.setMinimumSize(1100, 700)
        self._archive_view: ArchiveView | None = None
        self._uncoded_view: ArchiveView | None = None
        self._workspace_view: WorkspaceView | None = None
        self._build_ui()
        self._build_menu()
        self._build_toolbar()
        self._update_status()
        self._restore_geometry()

        # Auto-refresh ogni 30 secondi
        self._timer = QTimer()
        self._timer.timeout.connect(self._auto_refresh)
        self._timer.start(30_000)

    def _settings(self) -> QSettings:
        return QSettings(APP_NAME, "MainWindow")

    def _restore_geometry(self):
        s = self._settings()
        geom = s.value("geometry")
        if geom:
            self.restoreGeometry(geom)
        state = s.value("windowState")
        if state:
            self.restoreState(state)
        tab = s.value("activeTab")
        if tab is not None:
            try:
                self.tabs.setCurrentIndex(int(tab))
            except (ValueError, TypeError):
                pass

    def closeEvent(self, event):
        s = self._settings()
        s.setValue("geometry",    self.saveGeometry())
        s.setValue("windowState", self.saveState())
        s.setValue("activeTab",   self.tabs.currentIndex())
        if self._archive_view:
            self._archive_view.save_layout()
        if self._uncoded_view:
            self._uncoded_view.save_layout()
        if self._workspace_view:
            self._workspace_view.save_layout()
        super().closeEvent(event)

    # ------------------------------------------------------------------
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Tab principale
        self.tabs = QTabWidget()
        self.tabs.setTabPosition(QTabWidget.TabPosition.West)

        self._archive_view   = ArchiveView(view_mode="archive")
        self._uncoded_view   = ArchiveView(view_mode="uncoded")
        self._workspace_view = WorkspaceView()

        self._archive_view.document_selected.connect(self._open_document)
        self._uncoded_view.document_selected.connect(self._open_document)

        self.tabs.addTab(self._archive_view,   "🗄️  Archivio CAD")
        self.tabs.addTab(self._uncoded_view,   "🗂️  Non codificati")
        self.tabs.addTab(self._workspace_view, "📁  Workspace")

        layout.addWidget(self.tabs)

        # Status bar
        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.lbl_profile_status = QLabel()
        self.lbl_user_status    = QLabel()
        self.lbl_db_status      = QLabel()
        self.status.addPermanentWidget(self.lbl_profile_status)
        self.status.addPermanentWidget(self.lbl_db_status)
        self.status.addPermanentWidget(self.lbl_user_status)

    # ------------------------------------------------------------------
    def _build_menu(self):
        mb = self.menuBar()

        # File
        m_file = mb.addMenu("File")

        act_new = QAction("Nuovo documento…", self)
        act_new.setShortcut(QKeySequence("Ctrl+N"))
        act_new.triggered.connect(self._new_document)
        m_file.addAction(act_new)

        m_file.addSeparator()

        act_refresh = QAction("Aggiorna archive", self)
        act_refresh.setShortcut(QKeySequence("F5"))
        act_refresh.triggered.connect(self._refresh_all)
        m_file.addAction(act_refresh)

        m_file.addSeparator()

        act_logout = QAction("Cambia utente…", self)
        act_logout.triggered.connect(self._logout)
        m_file.addAction(act_logout)

        act_exit = QAction("Esci", self)
        act_exit.setShortcut(QKeySequence("Alt+F4"))
        act_exit.triggered.connect(self.close)
        m_file.addAction(act_exit)

        # Strumenti
        m_tools = mb.addMenu("Strumenti")

        act_schema = QAction("Schema di Codifica…", self)
        act_schema.triggered.connect(self._open_coding_schema)
        m_tools.addAction(act_schema)

        act_coding = QAction("Macchine e Gruppi…", self)
        act_coding.triggered.connect(self._open_coding)
        m_tools.addAction(act_coding)

        act_users = QAction("Gestione utenti…", self)
        act_users.triggered.connect(self._open_users)
        m_tools.addAction(act_users)

        m_tools.addSeparator()

        act_sw_cfg = QAction("🔧 Configurazione SolidWorks…", self)
        act_sw_cfg.triggered.connect(self._open_sw_config)
        m_tools.addAction(act_sw_cfg)

        m_tools.addSeparator()

        act_profiles = QAction("📋 PDM Profile…", self)
        act_profiles.triggered.connect(self._open_profiles)
        m_tools.addAction(act_profiles)

        act_setup = QAction("Configurazione percorso rete…", self)
        act_setup.triggered.connect(self._open_setup)
        m_tools.addAction(act_setup)

        m_tools.addSeparator()

        act_backup = QAction("💾 Backup database…", self)
        act_backup.triggered.connect(self._backup_db)
        m_tools.addAction(act_backup)

        m_tools.addSeparator()

        act_regen_thumb = QAction("🖼️  Rigenera anteprime…", self)
        act_regen_thumb.setToolTip("Rigenera le thumbnail dei documenti archiviati tramite eDrawings")
        act_regen_thumb.triggered.connect(self._regen_thumbnails)
        m_tools.addAction(act_regen_thumb)

        # SolidWorks
        m_sw = mb.addMenu("SolidWorks")
        act_sw_checkout = QAction("Checkout da SW…", self)
        act_sw_checkout.triggered.connect(self._sw_checkout)
        m_sw.addAction(act_sw_checkout)

        act_sw_checkin = QAction("Check-in da SW…", self)
        act_sw_checkin.triggered.connect(self._sw_checkin)
        m_sw.addAction(act_sw_checkin)

        m_sw.addSeparator()

        act_import_asm = QAction("📦 Importa struttura ASM da SW…", self)
        act_import_asm.setToolTip(
            "Apre il wizard per importare massivamente la struttura di un assieme "
            "aperto in SolidWorks, con codifica automatica e copia in workspace."
        )
        act_import_asm.triggered.connect(self._import_asm_wizard)
        m_sw.addAction(act_import_asm)

        m_sw.addSeparator()
        act_macro_info = QAction("Informazioni macro SW", self)
        act_macro_info.triggered.connect(self._show_macro_info)
        m_sw.addAction(act_macro_info)

        # Help
        m_help = mb.addMenu("?")
        act_manual = QAction("Manuale…", self)
        act_manual.triggered.connect(self._open_manual)
        m_help.addAction(act_manual)
        m_help.addSeparator()
        act_about = QAction(f"Informazioni su {APP_NAME}", self)
        act_about.triggered.connect(self._about)
        m_help.addAction(act_about)

    # ------------------------------------------------------------------
    def _build_toolbar(self):
        tb = QToolBar("Azioni principali")
        tb.setMovable(False)
        self.addToolBar(tb)

        act_new = QAction("➕ Nuovo", self)
        act_new.setToolTip("Crea nuovo documento (Ctrl+N)")
        act_new.triggered.connect(self._new_document)
        tb.addAction(act_new)

        tb.addSeparator()

        act_co = QAction("📤 Checkout", self)
        act_co.setToolTip("Esegui checkout del documento selezionato")
        act_co.triggered.connect(self._toolbar_checkout)
        tb.addAction(act_co)

        act_ci = QAction("📥 Check-in", self)
        act_ci.setToolTip("Esegui check-in del documento selezionato")
        act_ci.triggered.connect(self._toolbar_checkin)
        tb.addAction(act_ci)

        tb.addSeparator()

        act_wf = QAction("🔄 Workflow", self)
        act_wf.setToolTip("Cambia stato workflow (richiede selezione codice)")
        act_wf.triggered.connect(self._toolbar_workflow)
        tb.addAction(act_wf)

        tb.addSeparator()

        act_ref = QAction("↻ Aggiorna", self)
        act_ref.triggered.connect(self._refresh_all)
        tb.addAction(act_ref)

    # ------------------------------------------------------------------
    def _update_status(self):
        if session.is_logged_in:
            user = session.user
            self.lbl_user_status.setText(
                f"  👤 {user['full_name']} ({user['role']})  "
            )
        else:
            self.lbl_user_status.setText("  Non connesso  ")

        if session.is_connected:
            self.lbl_db_status.setText(
                f"  🗄️  DB: {session.sp.db_file.name}  "
            )
            self.lbl_db_status.setStyleSheet("color:#a6e3a1;")
        else:
            self.lbl_db_status.setText("  ⚠️  Nessun DB  ")
            self.lbl_db_status.setStyleSheet("color:#f38ba8;")

        # Profilo attivo
        profile = session.profile_name or ""
        if profile:
            self.lbl_profile_status.setText(f"  📋 {profile}  ")
            self.lbl_profile_status.setStyleSheet("color:#89b4fa;")
        else:
            self.lbl_profile_status.setText("")

    # ------------------------------------------------------------------
    def _new_document(self):
        if not session.can("create"):
            QMessageBox.warning(
                self, "Permesso negato",
                "Non hai i permessi per creare documenti "
                f"(ruolo: {session.user['role']})"
            )
            return
        dlg = DocumentDialog(parent=self)
        dlg.exec()
        self._refresh_all()

    def _open_document(self, doc_id: int):
        dlg = DocumentDialog(document_id=doc_id, parent=self)
        dlg.exec()
        self._refresh_all()

    # ------------------------------------------------------------------
    def _active_archive_like_view(self) -> ArchiveView | None:
        """Ritorna la vista archivio attiva (Archivio CAD o Non codificati)."""
        w = self.tabs.currentWidget()
        if w is self._archive_view:
            return self._archive_view
        if w is self._uncoded_view:
            return self._uncoded_view
        return None

    def _toolbar_checkout(self):
        view = self._active_archive_like_view()
        if view:
            doc_id = view._selected_doc_id()
            if not doc_id:
                QMessageBox.information(
                    self, "Info", "Selezionare un documento nell'archivio"
                )
                return
            view._action_checkout(doc_id)

    def _toolbar_checkin(self):
        view = self._active_archive_like_view()
        if view:
            doc_id = view._selected_doc_id()
            if not doc_id:
                QMessageBox.information(
                    self, "Info", "Selezionare un documento nell'archivio"
                )
                return
            view._action_checkin(doc_id)

    def _toolbar_workflow(self):
        view = self._active_archive_like_view()
        if not view:
            return
        handled = view.action_workflow_toolbar()
        if not handled:
            QMessageBox.information(
                self, "Workflow",
                "Selezionare un nodo codice nell'archivio per aprire il workflow."
            )

    # ------------------------------------------------------------------
    def _refresh_all(self):
        if self._archive_view:
            self._archive_view.refresh()
        if self._uncoded_view:
            self._uncoded_view.refresh()
        if self._workspace_view:
            self._workspace_view.refresh()

    def _auto_refresh(self):
        try:
            self._refresh_all()
        except Exception:
            pass

    # ------------------------------------------------------------------
    def _logout(self):
        session.logout()
        self._update_status()
        from ui.login_dialog import LoginDialog
        dlg = LoginDialog(parent=self)
        if dlg.exec():
            self._update_status()
            self._refresh_all()
            if self._workspace_view:
                self._workspace_view.refresh()

    def _open_coding_schema(self):
        if not session.can("admin"):
            QMessageBox.warning(
                self, "Accesso negato",
                "Solo gli Amministratori possono modificare lo schema di codifica."
            )
            return
        from ui.coding_schema_dialog import CodingSchemaDialog
        dlg = CodingSchemaDialog(parent=self)
        dlg.exec()

    def _open_coding(self):
        from ui.coding_dialog import CodingDialog
        dlg = CodingDialog(parent=self)
        dlg.exec()

    def _open_users(self):
        if not session.can("admin"):
            QMessageBox.warning(
                self, "Permesso negato",
                "Solo gli Amministratori possono gestire gli utenti"
            )
            return
        from ui.users_dialog import UsersDialog
        dlg = UsersDialog(parent=self)
        dlg.exec()

    def _open_setup(self):
        from ui.setup_dialog import SetupDialog
        dlg = SetupDialog(parent=self)
        dlg.exec()

    def _open_sw_config(self):
        from ui.sw_config_dialog import SWConfigDialog
        dlg = SWConfigDialog(parent=self)
        dlg.exec()

    def _open_profiles(self):
        from ui.profile_dialog import ProfileDialog
        dlg = ProfileDialog(parent=self)
        dlg.profile_switched.connect(self._on_profile_switched)
        dlg.exec()

    def _on_profile_switched(self, profile_name: str):
        try:
            session.switch_profile(profile_name)
            self._update_status()
            self._refresh_all()
            if not session.is_logged_in:
                from ui.login_dialog import LoginDialog
                login = LoginDialog(parent=self)
                if login.exec():
                    self._update_status()
                    self._refresh_all()
        except Exception as e:
            QMessageBox.critical(
                self, "Errore",
                f"Cambio profilo fallito:\n{e}"
            )

    # ------------------------------------------------------------------
    def _sw_checkout(self):
        """Interazione con SolidWorks aperto per il checkout."""
        QMessageBox.information(
            self, "SolidWorks – Checkout",
            "Per fare il checkout da SolidWorks:\n\n"
            "1. Aprire il file .SLDPRT/.SLDASM/.SLDDRW in SolidWorks\n"
            "2. Eseguire la macro PDM_Integration.swb\n"
            "3. La macro effettuerà automaticamente il checkout\n\n"
            "Percorso macro: macros/PDM_Integration.swb"
        )

    def _sw_checkin(self):
        QMessageBox.information(
            self, "SolidWorks – Check-in",
            "Per fare il check-in da SolidWorks:\n\n"
            "1. Salvare il file in SolidWorks\n"
            "2. Eseguire la macro PDM_Integration.swb\n"
            "3. La macro effettuerà automaticamente il check-in"
        )

    def _show_macro_info(self):
        import os
        from pathlib import Path
        macro_path = Path(__file__).parent.parent / "macros" / "PDM_Integration.swb"
        QMessageBox.information(
            self, "Macro SolidWorks",
            f"Posizione macro:\n{macro_path}\n\n"
            "Per installare la macro in SolidWorks:\n"
            "• Strumenti → Macro → Esegui → Seleziona il file .swp\n"
            "• Oppure: Strumenti → Personalizza → Macro → Assegna a pulsante barra"
        )

    def _import_asm_wizard(self):
        """Lancia il wizard di importazione massiva ASM come processo separato."""
        import subprocess
        from pathlib import Path
        import sys
        script = Path(__file__).parent.parent / "macros" / "pdm_asm_import.py"
        venv_py = Path(__file__).parent.parent / ".venv" / "Scripts" / "pythonw.exe"
        python_exe = str(venv_py) if venv_py.exists() else sys.executable
        try:
            subprocess.Popen(
                [python_exe, str(script)],
                cwd=str(Path(__file__).parent.parent),
            )
        except Exception as e:
            QMessageBox.critical(self, "Errore", f"Impossibile avviare il wizard:\n{e}")

    def _backup_db(self):
        if not session.is_connected:
            QMessageBox.warning(self, "Backup", "Nessun database connesso.")
            return
        from core.backup_manager import BackupManager
        bm = BackupManager(session.sp.db_file)
        try:
            dest = bm.create()
            backups = bm.list_backups()
            QMessageBox.information(
                self, "Backup completato",
                f"Backup creato:\n{dest}\n\n"
                f"Backup totali conservati: {len(backups)} (max {bm.keep})"
            )
        except Exception as e:
            QMessageBox.critical(self, "Errore backup", str(e))

    def _regen_thumbnails(self):
        if not session.is_connected:
            QMessageBox.warning(self, "Non connesso", "Connettersi prima al database.")
            return
        from ui.regen_thumbnails_dialog import RegenThumbnailsDialog
        dlg = RegenThumbnailsDialog(parent=self)
        dlg.exec()

    def _open_manual(self):
        from ui.manual_dialog import ManualDialog
        dlg = ManualDialog(parent=self)
        dlg.exec()

    def _about(self):
        QMessageBox.about(
            self, f"Informazioni su {APP_NAME}",
            f"<h3>{APP_NAME}  v{APP_VERSION}</h3>"
            "<p>Sistema PDM per la gestione di documenti SolidWorks</p>"
            "<ul>"
            "<li>Archivio CAD centralizzato su rete</li>"
            "<li>Check-in / Check-out con lock</li>"
            "<li>Gestione workflow documenti</li>"
            "<li>Codifica automatica</li>"
            "<li>Integrazione SolidWorks via COM / Macro</li>"
            "</ul>"
            "<p>Nessun server richiesto – max 5 utenti simultanei</p>"
        )

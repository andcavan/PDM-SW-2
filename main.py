#!/usr/bin/env python
# =============================================================================
#  main.py  –  Entry point PDM-SW
# =============================================================================
import sys
import os

# Assicura che la directory del progetto sia nel path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PyQt6.QtWidgets import QApplication, QMessageBox
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont

from config import (
    APP_NAME, load_local_config,
    get_profile_names, get_active_profile_name, set_active_profile,
)
from ui.styles import DARK_THEME
from ui.session import session


def main():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setStyle("Fusion")
    app.setStyleSheet(DARK_THEME)

    # Font globale
    font = QFont("Segoe UI", 10)
    app.setFont(font)

    # ----------------------------------------------------------------
    # 1) Selezione profilo
    # ----------------------------------------------------------------
    profiles = get_profile_names()

    if not profiles:
        # Nessun profilo → primo setup (crea profilo "Default")
        from ui.setup_dialog import SetupDialog
        dlg = SetupDialog()
        if dlg.exec() != dlg.DialogCode.Accepted:
            sys.exit(0)
        profiles = get_profile_names()

    if len(profiles) > 1:
        # Più profili → mostra selettore
        from ui.profile_dialog import ProfileSelector
        active = get_active_profile_name()
        selector = ProfileSelector(profiles, active)
        if selector.exec() != selector.DialogCode.Accepted:
            sys.exit(0)
        selected = selector.selected_profile
        if selected != get_active_profile_name():
            set_active_profile(selected)

    # ----------------------------------------------------------------
    # 2) Connessione al database (profilo attivo)
    # ----------------------------------------------------------------
    cfg = load_local_config()
    shared_root = cfg.get("shared_root", "")

    if not shared_root:
        QMessageBox.critical(
            None, "Errore",
            "Profilo attivo senza percorso condiviso configurato.\n"
            "Riconfigurare il percorso."
        )
        from ui.setup_dialog import SetupDialog
        dlg = SetupDialog()
        if dlg.exec() != dlg.DialogCode.Accepted:
            sys.exit(1)
        cfg = load_local_config()
        shared_root = cfg.get("shared_root", "")

    try:
        session.connect(shared_root)
        session.profile_name = get_active_profile_name()
    except Exception as e:
        QMessageBox.critical(
            None, "Errore di connessione",
            f"Impossibile connettersi al database PDM:\n\n{e}\n\n"
            f"Verificare che il percorso '{shared_root}' sia accessibile."
        )
        from ui.setup_dialog import SetupDialog
        dlg = SetupDialog()
        if dlg.exec() != dlg.DialogCode.Accepted:
            sys.exit(1)
        try:
            cfg = load_local_config()
            session.connect(cfg.get("shared_root", ""))
            session.profile_name = get_active_profile_name()
        except Exception as e2:
            QMessageBox.critical(None, "Errore fatale", str(e2))
            sys.exit(1)

    # ----------------------------------------------------------------
    # 3) Login utente
    # ----------------------------------------------------------------
    from ui.login_dialog import LoginDialog
    login_dlg = LoginDialog()
    if login_dlg.exec() != login_dlg.DialogCode.Accepted:
        sys.exit(0)

    # ----------------------------------------------------------------
    # 4) Finestra principale
    # ----------------------------------------------------------------
    from ui.main_window import MainWindow
    win = MainWindow()
    win.show()
    win._refresh_all()
    if win._workspace_view:
        win._workspace_view.refresh()
    win._update_status()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()

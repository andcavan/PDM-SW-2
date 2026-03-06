#!/usr/bin/env python
# =============================================================================
#  macros/pdm_asm_import.py  –  Wizard importazione massiva ASM da SolidWorks
#
#  Lanciato dalla macro VBA PDM_ImportAsm:
#    pythonw pdm_asm_import.py "C:\...\NOMEASM.SLDASM"
#
#  Oppure da menu SolidWorks nell'app principale.
# =============================================================================
from __future__ import annotations
import sys
import logging
import socket
from pathlib import Path

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

LOG_FILE = ROOT / "macros" / "sw_bridge.log"

def _setup_log():
    logging.basicConfig(
        filename=str(LOG_FILE),
        level=logging.DEBUG,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

_setup_log()


def _load_session():
    """Carica db, sp, user e i manager necessari."""
    from config import load_local_config, SharedPaths
    from core.database import Database
    from core.coding_manager import CodingManager
    from core.file_manager import FileManager
    from core.asm_manager import AsmManager

    cfg = load_local_config()
    shared_root = cfg.get("shared_root", "")
    if not shared_root:
        raise ValueError(
            "Percorso condiviso non configurato.\n"
            "Aprire PDM-SW e configurare il percorso di rete."
        )

    sp = SharedPaths(shared_root)
    db = Database(sp.db_file, sp.db_lock_file)

    hostname = socket.gethostname()
    user = (
        db.fetchone("SELECT * FROM users WHERE workstation=? AND active=1", (hostname,))
        or db.fetchone("SELECT * FROM users WHERE role='Amministratore' AND active=1")
        or db.fetchone("SELECT * FROM users WHERE active=1")
    )
    if not user:
        raise ValueError("Nessun utente trovato nel database PDM.")

    coding = CodingManager(db)
    files  = FileManager(db, sp, user)
    asm    = AsmManager(db)

    return db, sp, user, coding, files, asm


def main():
    asm_file = sys.argv[1] if len(sys.argv) > 1 else ""
    logging.info("pdm_asm_import avviato  |  file=%s", asm_file)

    from PyQt6.QtWidgets import QApplication, QMessageBox
    app = QApplication(sys.argv)
    app.setApplicationName("PDM-SW  –  Importazione ASM")
    app.setStyle("Fusion")

    try:
        db, sp, user, coding, files, asm = _load_session()
    except Exception as e:
        logging.error("Sessione fallita: %s", e, exc_info=True)
        QMessageBox.critical(None, "PDM-SW – Errore sessione", str(e))
        sys.exit(1)

    from ui.asm_import_wizard import AsmImportWizard
    dlg = AsmImportWizard(
        asm_file=Path(asm_file) if asm_file else None,
        db=db, sp=sp, user=user,
        coding=coding, files=files, asm_mgr=asm,
    )
    dlg.show()
    dlg.raise_()
    dlg.activateWindow()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

#!/usr/bin/env python
# =============================================================================
#  macros/sw_bridge.py  -  Bridge Python chiamato dalla macro SolidWorks  v2.2
#
#  Uso dalla macro VBA:
#    pythonw sw_bridge.py --action checkout --file "C:\...\CODICE.SLDPRT"
#    pythonw sw_bridge.py --action checkin  --file "C:\...\CODICE.SLDPRT"
#    pythonw sw_bridge.py --action undo_checkout --file "C:\...\CODICE.SLDPRT"
#    pythonw sw_bridge.py --action open
#    pythonw sw_bridge.py --action import_props --file "..."
#    pythonw sw_bridge.py --action export_props --file "..."
# =============================================================================
import sys
import os
import argparse
import logging
import traceback
from pathlib import Path
from datetime import datetime

# Aggiungi la directory radice del progetto al path
ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

# ---------------------------------------------------------------------------
#  Logging su file  (fondamentale per pythonw che non ha console)
# ---------------------------------------------------------------------------
LOG_FILE = ROOT / "macros" / "sw_bridge.log"
RESULT_FILE = ROOT / "macros" / "sw_bridge_result.txt"

def _setup_logging():
    """Configura logging su file per catturare errori di pythonw."""
    logging.basicConfig(
        filename=str(LOG_FILE),
        level=logging.DEBUG,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    # Limita dimensione log: se > 500KB, tronca
    try:
        if LOG_FILE.exists() and LOG_FILE.stat().st_size > 512_000:
            lines = LOG_FILE.read_text(encoding="utf-8", errors="replace").splitlines()
            LOG_FILE.write_text("\n".join(lines[-200:]) + "\n", encoding="utf-8")
    except Exception:
        pass

_setup_logging()
logging.info("=" * 60)
logging.info("sw_bridge avviato  |  argv=%s", sys.argv)
logging.info("ROOT=%s  |  cwd=%s", ROOT, os.getcwd())
logging.info("Python=%s", sys.executable)


def load_session():
    """Carica configurazione e connette al DB. Ritorna (db, sp, user)."""
    logging.info("load_session() inizio")
    from config import load_local_config, SharedPaths
    from core.database import Database

    cfg = load_local_config()
    shared_root = cfg.get("shared_root", "")
    logging.info("shared_root=%s", shared_root)
    if not shared_root:
        raise ValueError(
            "Percorso condiviso non configurato.\n"
            "Aprire PDM-SW e configurare il percorso."
        )

    sp = SharedPaths(shared_root)
    db = Database(sp.db_file, sp.db_lock_file)

    # Utente: cerca per hostname corrente
    import socket
    hostname = socket.gethostname()
    user_row = db.fetchone(
        "SELECT * FROM users WHERE workstation=? AND active=1", (hostname,)
    )
    if not user_row:
        user_row = db.fetchone(
            "SELECT * FROM users WHERE role='Amministratore' AND active=1"
        )
    if not user_row:
        user_row = db.fetchone("SELECT * FROM users WHERE active=1")
    if not user_row:
        raise ValueError("Nessun utente trovato nel database PDM.")

    logging.info("Sessione OK  user=%s (id=%s)", user_row.get("name"), user_row.get("id"))
    return db, sp, user_row


def find_document(db, file_path: str, user=None, need_lock: bool = False):
    """
    Trova il documento PDM dal percorso file workspace.
    Il file e' nominato {code}.SLDPRT/ASM/DRW.
    Ritorna il dict del documento o None.
    """
    from config import SW_EXTENSIONS

    fp = Path(file_path)
    code = fp.stem                     # es. "AAA_0100-0003"
    ext = fp.suffix.upper()            # es. ".SLDPRT"
    doc_type = SW_EXTENSIONS.get(ext, "")

    if need_lock and user:
        # Per checkin/undo: cerca quello in checkout dall'utente corrente
        doc = db.fetchone(
            "SELECT * FROM documents WHERE code=? AND doc_type=? "
            "AND is_locked=1 AND locked_by=? "
            "ORDER BY revision DESC",
            (code, doc_type, user["id"]),
        )
        if doc:
            return doc

    # Cerca per codice + tipo, revisione piu recente non obsoleta
    doc = db.fetchone(
        "SELECT * FROM documents WHERE code=? AND doc_type=? "
        "AND state != 'Obsoleto' "
        "ORDER BY revision DESC",
        (code, doc_type),
    )
    if doc:
        return doc

    # Fallback: cerca per file_name esatto (retrocompatibilita)
    doc = db.fetchone(
        "SELECT * FROM documents WHERE file_name=?",
        (fp.name,),
    )
    return doc


# ===========================================================================
#  AZIONI
# ===========================================================================

def action_checkout(file_path: str):
    db, sp, user = load_session()
    from core.checkout_manager import CheckoutManager
    co = CheckoutManager(db, sp, user)

    doc = find_document(db, file_path)
    if not doc:
        write_result(
            f"File '{Path(file_path).name}' non trovato nel PDM.\n"
            "Importare prima il documento nell'archivio PDM.",
            error=True,
        )
        return

    try:
        dest = co.checkout(doc["id"])
        write_result(f"Checkout eseguito.\nFile: {dest.name}")
    except PermissionError as e:
        write_result(str(e), error=True)
    except Exception as e:
        write_result(f"Errore checkout:\n{e}", error=True)


def action_checkin(file_path: str):
    db, sp, user = load_session()
    from core.checkout_manager import CheckoutManager
    from core.properties_manager import PropertiesManager
    from core.asm_manager import AsmManager
    co = CheckoutManager(db, sp, user)
    pm = PropertiesManager(db)
    asm = AsmManager(db)

    doc = find_document(db, file_path, user=user, need_lock=True)
    if not doc:
        write_result(
            f"File '{Path(file_path).name}' non e' in checkout da te\n"
            "o non e' stato trovato nel PDM.",
            error=True,
        )
        return

    # ------------------------------------------------------------------
    #  PRE-CHECKIN: importa proprietà e BOM dal file aperto in SolidWorks
    # ------------------------------------------------------------------
    sync_errors = []

    # 1) Importa proprietà dal file
    try:
        logging.info("Pre-checkin: lettura proprieta' da SW per doc %s", doc["code"])
        props = pm.read_from_sw_file(Path(file_path), file_name=Path(file_path).name)
        err = props.pop("_error", None)
        if err:
            sync_errors.append(f"Proprieta': {err}")
        elif props:
            pm.save_properties(doc["id"], props)
            logging.info("Pre-checkin: %d proprieta' salvate", len(props))
        else:
            logging.info("Pre-checkin: nessuna proprieta' custom nel file")
    except Exception as e:
        sync_errors.append(f"Proprieta': {e}")
        logging.error("Pre-checkin proprieta' fallito: %s", e, exc_info=True)

    # 2) Per gli assiemi, aggiorna la BOM
    if doc["doc_type"] == "Assieme":
        try:
            logging.info("Pre-checkin: aggiornamento BOM per assieme %s", doc["code"])
            n = asm.import_bom_from_active_doc(doc["id"])
            logging.info("Pre-checkin: %d componenti BOM aggiornati", n)
        except Exception as e:
            sync_errors.append(f"BOM: {e}")
            logging.error("Pre-checkin BOM fallito: %s", e, exc_info=True)

    # Se ci sono errori nell'importazione, avvisa e blocca il checkin
    if sync_errors:
        msg = "Check-in BLOCCATO.\n\n"
        msg += "Impossibile sincronizzare i dati da SolidWorks:\n\n"
        msg += "\n".join(f"  - {e}" for e in sync_errors)
        msg += "\n\nAssicurarsi che il file sia aperto in SolidWorks\n"
        msg += "e riprovare il check-in."
        write_result(msg, error=True)
        return

    # ------------------------------------------------------------------
    #  CHECKIN effettivo
    # ------------------------------------------------------------------
    try:
        result = co.checkin(
            doc["id"],
            archive_file=True,
            delete_from_workspace=False,
        )
        msg = "Check-in eseguito.\n"
        if result.get("archived"):
            msg += "File archiviato nel PDM.\n"
        else:
            msg += "Lock rilasciato (file non archiviato).\n"
        if result.get("conflict"):
            msg += "\nATTENZIONE: conflitto rilevato!\n" \
                   "Il file in archivio era stato modificato.\n"

        # Riepilogo sync
        msg += f"\nProprieta' importate: {len(props) if props else 0}"
        if doc["doc_type"] == "Assieme":
            msg += f"\nComponenti BOM aggiornati: {n}"

        write_result(msg)
    except PermissionError as e:
        write_result(str(e), error=True)
    except Exception as e:
        write_result(f"Errore check-in:\n{e}", error=True)


def action_undo_checkout(file_path: str):
    db, sp, user = load_session()
    from core.checkout_manager import CheckoutManager
    co = CheckoutManager(db, sp, user)

    doc = find_document(db, file_path, user=user, need_lock=True)
    if not doc:
        write_result(
            f"File '{Path(file_path).name}' non e' in checkout da te.",
            error=True,
        )
        return

    try:
        co.undo_checkout(doc["id"], delete_from_workspace=False)
        write_result("Checkout annullato.\nLock rilasciato senza archiviare.")
    except PermissionError as e:
        write_result(str(e), error=True)
    except Exception as e:
        write_result(f"Errore annulla checkout:\n{e}", error=True)


def action_import_props(file_path: str):
    """Legge le proprieta da SolidWorks e le salva nel PDM."""
    db, sp, user = load_session()
    from core.properties_manager import PropertiesManager

    doc = find_document(db, file_path)
    if not doc:
        write_result(
            f"File '{Path(file_path).name}' non trovato nel PDM.",
            error=True,
        )
        return

    pm = PropertiesManager(db)
    try:
        logging.info("Lettura proprieta da SW: %s", file_path)
        props = pm.read_from_sw_file(Path(file_path))

        # Controlla errore di lettura
        err = props.pop("_error", None)
        if err:
            write_result(f"Errore lettura proprieta SW:\n{err}", error=True)
            return

        if not props:
            write_result(
                "Nessuna proprieta' custom trovata nel file SolidWorks.",
                error=True,
            )
            return

        logging.info("Proprieta lette: %s", list(props.keys()))
        pm.save_properties(doc["id"], props)
        write_result(f"{len(props)} proprieta importate da SolidWorks nel PDM.")
    except Exception as e:
        logging.error("Errore import_props: %s", e, exc_info=True)
        write_result(str(e), error=True)


def action_export_props(file_path: str):
    """Legge le proprieta dal PDM e le scrive nel file SolidWorks."""
    db, sp, user = load_session()
    from core.properties_manager import PropertiesManager

    doc = find_document(db, file_path)
    if not doc:
        write_result(
            f"File '{Path(file_path).name}' non trovato nel PDM.",
            error=True,
        )
        return

    pm = PropertiesManager(db)
    try:
        props = pm.get_properties(doc["id"])
        if not props:
            write_result("Nessuna proprieta' nel PDM per questo documento.", error=True)
            return
        logging.info("Export proprieta verso SW: %s", list(props.keys()))
        pm.write_to_sw_file(Path(file_path), props)
        write_result(f"{len(props)} proprieta esportate dal PDM in SolidWorks.")
    except Exception as e:
        logging.error("Errore export_props: %s", e, exc_info=True)
        write_result(str(e), error=True)


def action_open():
    """Apre l'interfaccia principale PDM-SW."""
    import subprocess
    script = ROOT / "main.py"
    # Usa il python del venv se disponibile
    venv_python = ROOT / ".venv" / "Scripts" / "pythonw.exe"
    python = str(venv_python) if venv_python.exists() else sys.executable
    logging.info("action_open: lancio %s %s", python, script)
    subprocess.Popen([python, str(script)], cwd=str(ROOT))
    write_result("Applicazione PDM-SW avviata.")


def action_import_props_json(file_path: str):
    """
    Importa proprieta' dal file JSON scritto dalla macro VBA.
    La VBA legge le proprieta' direttamente da SolidWorks (piu' affidabile
    che usare COM da un processo Python separato).
    """
    db, sp, user = load_session()
    from core.properties_manager import PropertiesManager
    import json

    doc = find_document(db, file_path)
    if not doc:
        write_result(
            f"File '{Path(file_path).name}' non trovato nel PDM.",
            error=True,
        )
        return

    props_file = ROOT / "macros" / "sw_props_temp.json"
    if not props_file.exists():
        write_result(
            "File proprieta' temporaneo non trovato.\n"
            "La macro VBA non ha scritto le proprieta'.",
            error=True,
        )
        return

    try:
        raw = props_file.read_text(encoding="utf-8-sig", errors="replace")
        logging.info("JSON props raw: %s", raw[:500])
        props = json.loads(raw)
    except Exception as e:
        write_result(f"Errore lettura JSON proprieta':\n{e}", error=True)
        return
    finally:
        # Pulisci file temporaneo
        try:
            props_file.unlink(missing_ok=True)
        except Exception:
            pass

    if not props:
        write_result("Nessuna proprieta' trovata nel file.", error=True)
        return

    pm = PropertiesManager(db)
    logging.info("Salvataggio %d proprieta': %s", len(props), list(props.keys()))
    pm.save_properties(doc["id"], props)
    write_result(f"{len(props)} proprieta' importate da SolidWorks nel PDM.")


def write_result(msg: str, error: bool = False):
    """
    Scrive il risultato in un file che la macro VBA leggera'.
    Formato: prima riga = 'OK' o 'ERR', resto = messaggio.
    NON usare MessageBox qui: il processo gira nascosto e bloccherebbe tutto.
    """
    level = "ERROR" if error else "INFO"
    logging.info("write_result [%s]: %s", level, msg)
    try:
        prefix = "ERR" if error else "OK"
        RESULT_FILE.write_text(f"{prefix}\n{msg}", encoding="utf-8")
    except Exception as e:
        logging.error("Impossibile scrivere result file: %s", e)


# ===========================================================================
def main():
    try:
        _main_inner()
    except SystemExit:
        raise
    except Exception:
        # Cattura QUALSIASI errore e scrivilo nel log
        tb = traceback.format_exc()
        logging.critical("ECCEZIONE NON GESTITA:\n%s", tb)
        try:
            write_result(f"Errore critico PDM-SW:\n{tb[:500]}", error=True)
        except Exception:
            pass
        sys.exit(1)


def _main_inner():
    parser = argparse.ArgumentParser(description="PDM-SW SolidWorks Bridge")
    parser.add_argument(
        "--action",
        choices=["checkout", "checkin", "undo_checkout",
                 "import_props", "import_props_json",
                 "export_props", "open"],
        required=True,
    )
    parser.add_argument("--file", default="",
                        help="Percorso file SolidWorks")
    args = parser.parse_args()

    logging.info("Azione: %s  |  File: %s", args.action, args.file)

    try:
        if args.action == "checkout":
            if not args.file:
                write_result("--file richiesto per checkout", error=True)
                sys.exit(1)
            action_checkout(args.file)

        elif args.action == "checkin":
            if not args.file:
                write_result("--file richiesto per checkin", error=True)
                sys.exit(1)
            action_checkin(args.file)

        elif args.action == "undo_checkout":
            if not args.file:
                write_result("--file richiesto per undo_checkout", error=True)
                sys.exit(1)
            action_undo_checkout(args.file)

        elif args.action == "import_props":
            if not args.file:
                write_result("--file richiesto per import_props", error=True)
                sys.exit(1)
            action_import_props(args.file)

        elif args.action == "import_props_json":
            if not args.file:
                write_result("--file richiesto per import_props_json", error=True)
                sys.exit(1)
            action_import_props_json(args.file)

        elif args.action == "export_props":
            if not args.file:
                write_result("--file richiesto per export_props", error=True)
                sys.exit(1)
            action_export_props(args.file)

        elif args.action == "open":
            action_open()

    except Exception as e:
        logging.error("Errore azione %s: %s", args.action, e, exc_info=True)
        write_result(f"Errore PDM-SW:\n{e}", error=True)
        sys.exit(1)

    logging.info("Azione %s completata con successo", args.action)


if __name__ == "__main__":
    main()

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

    logging.info("Sessione OK  user=%s (id=%s)", user_row.get("username"), user_row.get("id"))
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

    # 1) Importa proprieta dal file (custom + mapping SW->PDM)
    try:
        logging.info("Pre-checkin: lettura proprieta' da SW per doc %s", doc["code"])
        sync_in = pm.sync_sw_to_pdm(doc["id"], Path(file_path), file_name=Path(file_path).name)
        if not sync_in.get("ok"):
            sync_errors.append(f"Proprieta': {sync_in.get('error', '')}")
            props = {}
        elif int(sync_in.get("imported_count", 0)) > 0:
            props = dict(sync_in.get("props") or {})
            logging.info("Pre-checkin: %d proprieta' salvate", int(sync_in.get("imported_count", 0)))
        else:
            props = {}
            logging.info("Pre-checkin: nessuna proprieta' custom nel file")
    except Exception as e:
        sync_errors.append(f"Proprieta': {e}")
        logging.error("Pre-checkin proprieta' fallito: %s", e, exc_info=True)

    # 1b) Enforce revisione PDM -> SW prima del check-in
    try:
        sync_out = pm.sync_pdm_to_sw(doc["id"], Path(file_path), force_revision=True)
        if not sync_out.get("ok"):
            sync_errors.append(f"Revisione: {sync_out.get('error', '')}")
        else:
            logging.info("Pre-checkin: %d proprieta' forzate da PDM verso SW", int(sync_out.get("written_count", 0)))
    except Exception as e:
        sync_errors.append(f"Revisione: {e}")
        logging.error("Pre-checkin revisione PDM->SW fallito: %s", e, exc_info=True)

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

        # POST-CHECKIN: genera thumbnail via eDrawings (non bloccante)
        _generate_thumbnail(doc, sp)

        write_result(msg)
    except PermissionError as e:
        write_result(str(e), error=True)
    except Exception as e:
        write_result(f"Errore check-in:\n{e}", error=True)


def _generate_thumbnail(doc: dict, sp):
    """
    Genera thumbnail PNG del file archiviato tramite eDrawings.
    Non bloccante: errori vengono solo loggati senza impatto sul checkin.
    """
    try:
        archive_path = sp.archive_path(doc["code"], doc["revision"])
        file_name = doc.get("file_name")
        if not file_name:
            return
        src_file = archive_path / file_name
        if not src_file.exists():
            logging.info("Thumbnail: file archivio non trovato: %s", src_file)
            return

        sp.thumbnails.mkdir(parents=True, exist_ok=True)
        dest = sp.thumbnails / f"{doc['code']}_{doc['revision']}.png"

        import win32com.client
        import time

        logging.info("Thumbnail: avvio eDrawings per %s", src_file)
        eApp = win32com.client.Dispatch("EModelView.EModelViewControl")
        eApp.OpenDoc(str(src_file), False, False, True, "")

        # Attendi che il documento sia caricato
        for _ in range(30):
            time.sleep(0.5)
            try:
                if eApp.FileName:
                    break
            except Exception:
                pass

        eApp.SaveAs(str(dest))
        eApp.CloseActiveDoc("")
        logging.info("Thumbnail generata: %s", dest)

    except Exception as e:
        logging.warning("Thumbnail generation fallita (non bloccante): %s", e)


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
        sync = pm.sync_sw_to_pdm(doc["id"], Path(file_path), file_name=Path(file_path).name)
        if not sync.get("ok"):
            write_result(f"Errore lettura proprieta SW:\n{sync.get('error', '')}", error=True)
            return

        if int(sync.get("imported_count", 0)) <= 0:
            if bool(sync.get("updated_owner", False)):
                write_result(
                    "Nessuna proprieta' custom trovata, ma i campi PDM mappati sono stati aggiornati.",
                    error=False,
                )
            else:
                write_result(
                    "Nessuna proprieta' custom trovata nel file SolidWorks.",
                    error=False,
                )
            return

        props = dict(sync.get("props") or {})
        logging.info("Proprieta lette: %s", list(props.keys()))
        write_result(f"{int(sync.get('imported_count', 0))} proprieta importate da SolidWorks nel PDM.")
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
        # 1) Sync mapping PDM->SW (revision/title/description/code/state/created_*).
        sync = pm.sync_pdm_to_sw(doc["id"], Path(file_path), force_revision=True)
        if not sync.get("ok"):
            write_result(f"Errore sync PDM->SW:\n{sync.get('error', '')}", error=True)
            return

        written = int(sync.get("written_count", 0))
        mapped_keys = set((sync.get("props") or {}).keys())

        # 2) Export custom properties presenti nella tabella PDM,
        #    escludendo le chiavi già scritte dalla mappatura per evitare
        #    una seconda scrittura (e secondo Save) sugli stessi campi.
        all_props = pm.get_properties(doc["id"])
        extra_props = {k: v for k, v in all_props.items() if k not in mapped_keys}
        if extra_props:
            logging.info("Export custom proprieta verso SW: %s", list(extra_props.keys()))
            pm.write_to_sw_file(Path(file_path), extra_props)

        if written <= 0 and not all_props:
            write_result("Nessuna proprieta' da esportare verso SolidWorks.", error=False)
            return

        msg = f"Sync mapping PDM->SW completata ({written} campi)."
        if extra_props:
            msg += f"\nCustom properties aggiuntive esportate: {len(extra_props)}."
        elif all_props:
            msg += f"\nCustom properties gia' coperte dalla mappatura: {len(all_props)}."
        write_result(msg, error=False)
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


def action_panel(file_path: str = ""):
    """
    Apre il pannello Qt PDM (pdm_panel.py) in background.
    La VBA NON aspetta la chiusura del pannello.
    """
    import subprocess
    panel_script = ROOT / "macros" / "pdm_panel.py"
    venv_python  = ROOT / ".venv" / "Scripts" / "pythonw.exe"
    python       = str(venv_python) if venv_python.exists() else sys.executable

    args = [python, str(panel_script)]
    if file_path:
        args.append(file_path)

    logging.info("action_panel: lancio %s", args)
    subprocess.Popen(args, cwd=str(ROOT))
    write_result("Pannello PDM avviato.")


def action_create_code(file_path: str):
    """
    Legge i parametri da macros/create_code_params.json scritto dalla VBA,
    assegna il codice via CodingManager, crea il documento nel DB.
    Ritorna il codice assegnato nel result file.
    """
    import json
    db, sp, user = load_session()
    from core.coding_manager import CodingManager
    from core.file_manager import FileManager

    params_file = ROOT / "macros" / "create_code_params.json"
    if not params_file.exists():
        write_result("File parametri 'create_code_params.json' non trovato.", error=True)
        return

    try:
        params = json.loads(params_file.read_text(encoding="utf-8-sig"))
    except Exception as e:
        write_result(f"Errore lettura parametri JSON:\n{e}", error=True)
        return
    finally:
        try:
            params_file.unlink(missing_ok=True)
        except Exception:
            pass

    title      = params.get("title", "")
    machine_id = params.get("machine_id")
    group_id   = params.get("group_id")
    subtype    = params.get("subtype", "PRT")    # PRT / ASM_MACH / ASM_GRP / ASM_SUB
    doc_level  = params.get("doc_level", 2)

    if not title:
        write_result("Campo 'title' obbligatorio nei parametri.", error=True)
        return
    if machine_id is None:
        write_result("Campo 'machine_id' obbligatorio nei parametri.", error=True)
        return

    coding = CodingManager(db)
    files  = FileManager(db, sp, user)

    try:
        if subtype == "ASM_MACH":
            code = coding.next_code_liv0(machine_id)
        elif subtype == "ASM_GRP":
            code = coding.next_code_liv1(machine_id, group_id)
        elif subtype == "ASM_SUB":
            code = coding.next_code_liv2_part(machine_id, group_id)
        else:  # PRT
            code = coding.next_code_liv2_subgroup(machine_id, group_id)

        fp = Path(file_path)
        from config import SW_EXTENSIONS
        doc_type = SW_EXTENSIONS.get(fp.suffix.upper(), "Parte")

        doc_id = files.create_document(
            code=code, revision="00", doc_type=doc_type, title=title,
            machine_id=machine_id, group_id=group_id, doc_level=doc_level,
        )
        logging.info("create_code: codice=%s id=%d", code, doc_id)
        write_result(f"Documento creato:\nCodice: {code}  rev.00  [{doc_type}]\n{title}")

    except Exception as e:
        logging.error("Errore create_code: %s", e, exc_info=True)
        write_result(f"Errore generazione codice:\n{e}", error=True)


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
                 "export_props", "open",
                 "panel", "create_code"],
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

        elif args.action == "panel":
            action_panel(args.file)

        elif args.action == "create_code":
            if not args.file:
                write_result("--file richiesto per create_code", error=True)
                sys.exit(1)
            action_create_code(args.file)

    except Exception as e:
        logging.error("Errore azione %s: %s", args.action, e, exc_info=True)
        write_result(f"Errore PDM-SW:\n{e}", error=True)
        sys.exit(1)

    logging.info("Azione %s completata con successo", args.action)


if __name__ == "__main__":
    main()

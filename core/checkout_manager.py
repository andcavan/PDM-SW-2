# =============================================================================
#  core/checkout_manager.py  –  Gestione check-in / check-out  v2.0
# =============================================================================
from __future__ import annotations
import hashlib
import os
import shutil
import socket
from pathlib import Path
from typing import Optional, TYPE_CHECKING

if TYPE_CHECKING:
    from core.database import Database
    from config import SharedPaths

# Stati che impediscono checkout / checkin (sola consultazione)
# Rilasciato e Obsoleto sono documenti ufficiali: non possono essere
# modificati né archiviati.
READONLY_STATES = ("Rilasciato", "Obsoleto")


class CheckoutManager:
    def __init__(self, db: "Database", shared_paths: "SharedPaths",
                 current_user: dict):
        self.db           = db
        self.sp           = shared_paths
        self.current_user = current_user

    # ==================================================================
    #  Utility
    # ==================================================================
    @staticmethod
    def _set_readonly(path: Path):
        """Imposta il file come sola lettura (tutti i permessi di scrittura rimossi)."""
        import stat
        try:
            path.chmod(path.stat().st_mode & ~(stat.S_IWRITE | stat.S_IWGRP | stat.S_IWOTH))
        except OSError:
            pass

    @staticmethod
    def _set_writable(path: Path):
        """Rimuove la sola lettura dal file (necessario su Windows prima di unlink/modifica)."""
        import stat
        try:
            path.chmod(path.stat().st_mode | stat.S_IWRITE)
        except OSError:
            pass

    @staticmethod
    def _md5(filepath: Path) -> str:
        """Calcola MD5 di un file."""
        h = hashlib.md5()
        with open(filepath, "rb") as f:
            for chunk in iter(lambda: f.read(8192), b""):
                h.update(chunk)
        return h.hexdigest()

    @staticmethod
    def _file_snapshot(filepath: Path) -> dict:
        """Ritorna snapshot {md5, size, mtime} di un file."""
        if not filepath or not filepath.exists():
            return {"md5": "", "size": 0, "mtime": 0.0}
        stat = filepath.stat()
        return {
            "md5":   CheckoutManager._md5(filepath),
            "size":  stat.st_size,
            "mtime": stat.st_mtime,
        }

    def _archive_file_path(self, doc: dict) -> Optional[Path]:
        """Percorso fisico del file in archivio."""
        if doc.get("archive_path"):
            return self.sp.root / doc["archive_path"]
        if doc.get("file_name"):
            p = self.sp.archive_path(doc["code"], doc["revision"])
            return p / doc["file_name"]
        return None

    def _ws_file_path(self, doc: dict) -> Path:
        """Percorso atteso del file nella workspace locale configurata."""
        from config import load_local_config
        from core.file_manager import EXT_FOR_TYPE
        cfg = load_local_config()
        ws = Path(cfg.get("sw_workspace", ""))
        ext = EXT_FOR_TYPE.get(doc.get("doc_type", ""), ".SLDPRT")
        return ws / (doc["code"] + ext)

    def _validate_state_for_edit(self, doc: dict, action: str = "checkout"):
        """Solleva PermissionError se lo stato non consente checkout/checkin."""
        if doc["state"] in READONLY_STATES:
            raise PermissionError(
                f"Impossibile fare {action}: il documento è '{doc['state']}'.\n"
                "Per modificare, creare una nuova revisione."
            )

    # ==================================================================
    #  CHECKOUT singolo
    # ==================================================================
    def checkout(self, document_id: int,
                 include_drw: bool = False) -> Path:
        """
        Checkout di un singolo documento:
        - Verifica stato e lock
        - Copia archivio → workspace
        - Lock + snapshot MD5
        - Registra in workspace_files
        Se include_drw=True, copia anche il DRW associato.
        Ritorna il percorso del file principale in workspace.
        """
        doc = self._get_doc_or_raise(document_id)
        self._validate_state_for_edit(doc)
        self._check_not_locked(doc)

        # Copia file in workspace (arriva come sola lettura da _copy_archive_to_workspace)
        dest = self._copy_archive_to_workspace(doc)
        # Il file è mio checkout: rendo scrivibile
        self._set_writable(dest)

        # Snapshot dell'archivio
        archive_file = self._archive_file_path(doc)
        snap = self._file_snapshot(archive_file) if archive_file else \
               {"md5": "", "size": 0, "mtime": 0.0}

        # Lock nel DB
        uid = self.current_user["id"]
        self.db.execute(
            """UPDATE documents
               SET is_locked=1, locked_by=?, locked_at=datetime('now'),
                   locked_ws=?, checkout_md5=?, checkout_size=?, checkout_mtime=?
               WHERE id=?""",
            (uid, socket.gethostname(),
             snap["md5"], snap["size"], snap["mtime"],
             document_id),
        )

        # Log
        self.db.execute(
            """INSERT INTO checkout_log
               (document_id, user_id, action, workstation, workspace_path,
                checkout_md5, checkout_size, checkout_mtime)
               VALUES (?,?,'checkout',?,?,?,?,?)""",
            (document_id, uid, socket.gethostname(), str(dest),
             snap["md5"], snap["size"], snap["mtime"]),
        )

        # Registra in workspace_files
        self._register_workspace_file(document_id, uid, "checkout", str(dest))

        # DRW opzionale
        if include_drw:
            self._checkout_companion_drw(doc, uid)

        return dest

    # ==================================================================
    #  CHECKOUT ASM (ricorsivo)
    # ==================================================================
    def checkout_asm(self, document_id: int,
                     include_drw: bool = False) -> dict:
        """
        Checkout di un assieme:
        - Checkout (lock) sull'ASM stesso
        - Copia tutti i componenti nella workspace SENZA lock (role='component')
        - Se include_drw, copia anche DRW dell'ASM (con lock)
        Ritorna {asm_path, component_paths: [...], drw_path}.
        """
        from core.asm_manager import AsmManager

        doc = self._get_doc_or_raise(document_id)
        if doc["doc_type"] != "Assieme":
            raise ValueError("checkout_asm è disponibile solo per documenti di tipo Assieme")

        # Checkout dell'ASM (con lock)
        asm_path = self.checkout(document_id, include_drw=include_drw)

        # Copia componenti senza lock
        asm_mgr = AsmManager(self.db)
        components = asm_mgr.get_bom_flat(document_id)
        uid = self.current_user["id"]
        component_paths = []

        for comp in components:
            child_doc = self._get_doc_or_raise(comp["child_id"])
            try:
                cp = self._copy_archive_to_workspace(child_doc)
                self._register_workspace_file(
                    comp["child_id"], uid, "component", str(cp),
                    parent_checkout_id=document_id
                )
                component_paths.append(cp)
            except Exception:
                # Se il componente non ha file archiviato, skip
                pass

        return {
            "asm_path": asm_path,
            "component_paths": component_paths,
            "drw_path": None,  # già gestito in checkout() se include_drw
        }

    # ==================================================================
    #  CHECKIN  (singolo documento)
    # ==================================================================
    def checkin(self, document_id: int,
                archive_file: bool = True,
                delete_from_workspace: bool = False,
                notes: str = "") -> dict:
        """
        Check-in di un singolo documento:
        - Confronto MD5 per rilevare modifiche / conflitti
        - Copia workspace → archivio (se archive_file=True)
        - Rilascio lock
        - Rimuove da workspace_files

        Ritorna dict: {modified: bool, conflict: bool, archived: bool, path: str}
        """
        doc = self._get_doc_or_raise(document_id)
        self._validate_state_for_edit(doc, "checkin")

        if not doc["is_locked"]:
            raise ValueError("Il documento non è in checkout")

        uid = self.current_user["id"]
        if doc["locked_by"] != uid:
            locker = self.db.fetchone(
                "SELECT full_name FROM users WHERE id=?", (doc["locked_by"],)
            )
            name = locker["full_name"] if locker else "altro utente"
            raise PermissionError(
                f"Impossibile fare checkin: il documento è bloccato da {name}"
            )

        # Percorso file workspace
        ws_file = self._ws_file_path(doc)
        result = {"modified": False, "conflict": False, "archived": False, "path": str(ws_file)}

        if ws_file.exists():
            ws_snap = self._file_snapshot(ws_file)

            # Confronto con snapshot al checkout (modifica locale?)
            checkout_md5 = doc.get("checkout_md5") or ""
            result["modified"] = (ws_snap["md5"] != checkout_md5) if checkout_md5 else True

            # Confronto con file attualmente in archivio (conflitto?)
            archive_file_path = self._archive_file_path(doc)
            if archive_file_path and archive_file_path.exists():
                archive_current_md5 = self._md5(archive_file_path)
                if checkout_md5 and archive_current_md5 != checkout_md5:
                    result["conflict"] = True

            # Archiviazione
            if archive_file:
                self._copy_workspace_to_archive(doc, ws_file)
                result["archived"] = True

            # Eliminazione dalla workspace o impostazione sola lettura
            if delete_from_workspace:
                try:
                    self._set_writable(ws_file)  # Windows: necessario prima di unlink
                    ws_file.unlink()
                except OSError:
                    pass
            else:
                # File rimane in workspace ma non è più in checkout: sola lettura
                self._set_readonly(ws_file)
        elif archive_file:
            # File non trovato in workspace ma si voleva archiviare → errore esplicito
            raise FileNotFoundError(
                f"File non trovato in workspace: {ws_file}\n"
                f"Impossibile archiviare {doc['code']} rev.{doc['revision']}."
            )

        # Rilascio lock
        self.db.execute(
            """UPDATE documents
               SET is_locked=0, locked_by=NULL, locked_at=NULL, locked_ws=NULL,
                   checkout_md5=NULL, checkout_size=NULL, checkout_mtime=NULL,
                   modified_by=?, modified_at=datetime('now')
               WHERE id=?""",
            (uid, document_id),
        )

        # Log
        self.db.execute(
            """INSERT INTO checkout_log
               (document_id, user_id, action, workstation, notes)
               VALUES (?,?,'checkin',?,?)""",
            (document_id, uid, socket.gethostname(), notes),
        )

        # Rimuovi da workspace_files
        self._unregister_workspace_file(document_id, uid)

        return result

    # ==================================================================
    #  CHECK FILE MODIFICATION
    # ==================================================================
    def is_file_modified(self, document_id: int) -> dict:
        """
        Verifica se il file in workspace è stato modificato rispetto allo snapshot.
        Ritorna {modified: bool, conflict: bool, ws_exists: bool}
        """
        doc = self._get_doc_or_raise(document_id)
        ws_file = self._ws_file_path(doc)

        if not ws_file.exists():
            return {"modified": False, "conflict": False, "ws_exists": False}

        ws_snap = self._file_snapshot(ws_file)
        checkout_md5 = doc.get("checkout_md5") or ""
        modified = (ws_snap["md5"] != checkout_md5) if checkout_md5 else True

        # Conflitto: archivio cambiato da quando abbiamo fatto checkout
        conflict = False
        archive_file = self._archive_file_path(doc)
        if archive_file and archive_file.exists() and checkout_md5:
            archive_md5 = self._md5(archive_file)
            conflict = (archive_md5 != checkout_md5)

        return {"modified": modified, "conflict": conflict, "ws_exists": True}

    # ==================================================================
    #  UNDO CHECKOUT (solo admin o stesso utente)
    # ==================================================================
    def undo_checkout(self, document_id: int,
                      delete_from_workspace: bool = False) -> bool:
        """
        Annulla il checkout: rilascia lock senza archiviare.
        Solo admin può annullare checkout di altri utenti.
        """
        doc = self._get_doc_or_raise(document_id)
        if not doc["is_locked"]:
            return False

        uid = self.current_user["id"]
        is_admin = self.current_user.get("role") == "Amministratore"

        if doc["locked_by"] != uid and not is_admin:
            raise PermissionError(
                "Solo l'amministratore può annullare il checkout di un altro utente."
            )

        # Rilascio lock
        self.db.execute(
            """UPDATE documents
               SET is_locked=0, locked_by=NULL, locked_at=NULL, locked_ws=NULL,
                   checkout_md5=NULL, checkout_size=NULL, checkout_mtime=NULL
               WHERE id=?""",
            (document_id,),
        )

        # Eliminazione file workspace se richiesto
        if delete_from_workspace:
            ws_file = self._ws_file_path(doc)
            if ws_file.exists():
                try:
                    self._set_writable(ws_file)  # Windows: necessario prima di unlink
                    ws_file.unlink()
                except OSError:
                    pass

        # Log
        self.db.execute(
            """INSERT INTO checkout_log
               (document_id, user_id, action, workstation, notes)
               VALUES (?,?,'undo_checkout',?,?)""",
            (document_id, uid, socket.gethostname(), "Checkout annullato"),
        )

        # Rimuovi da workspace_files (tutti i ruoli)
        self._unregister_workspace_file(document_id, doc.get("locked_by", uid))

        return True

    # ==================================================================
    #  CHECKOUT NUOVO DA WORKSPACE (creazione/importazione via SW)
    # ==================================================================
    def checkout_new_from_workspace(self, document_id: int, ws_file: Path) -> Path:
        """
        Registra un file appena creato/importato via SolidWorks come checked out.
        Il file in workspace (ws_file) è già pronto (da SaveAs SW).
        1. Copia ws_file → archivio con nome codice PDM
        2. Imposta archivio sola lettura
        3. Aggiorna DB: file_name, archive_path
        4. Imposta lock checkout (is_locked=1)
        5. Registra in workspace_files e checkout_log
        Ritorna il path del file in archivio.
        """
        doc = self._get_doc_or_raise(document_id)

        ext = self._ext_for_doc_type(doc["doc_type"])
        arch_dir = self.sp.archive_path(doc["code"], doc["revision"])
        arch_dir.mkdir(parents=True, exist_ok=True)
        arch_file = arch_dir / (doc["code"] + ext)

        if arch_file.exists():
            self._set_writable(arch_file)
        shutil.copy2(ws_file, arch_file)
        self._set_readonly(arch_file)  # archivio sempre sola lettura

        rel_path = str(arch_file.relative_to(self.sp.root))
        uid = self.current_user["id"]

        # Aggiorna metadati documento
        self.db.execute(
            """UPDATE documents
               SET file_name=?, file_ext=?, archive_path=?,
                   modified_by=?, modified_at=datetime('now')
               WHERE id=?""",
            (arch_file.name, ext, rel_path, uid, document_id),
        )

        # Snapshot del file archiviato
        snap = self._file_snapshot(arch_file)

        # Lock checkout
        self.db.execute(
            """UPDATE documents
               SET is_locked=1, locked_by=?, locked_at=datetime('now'),
                   locked_ws=?, checkout_md5=?, checkout_size=?, checkout_mtime=?
               WHERE id=?""",
            (uid, socket.gethostname(),
             snap["md5"], snap["size"], snap["mtime"], document_id),
        )

        # Log
        self.db.execute(
            """INSERT INTO checkout_log
               (document_id, user_id, action, workstation, workspace_path,
                checkout_md5, checkout_size, checkout_mtime)
               VALUES (?,?,'checkout',?,?,?,?,?)""",
            (document_id, uid, socket.gethostname(), str(ws_file),
             snap["md5"], snap["size"], snap["mtime"]),
        )

        # Registra workspace_files
        self._register_workspace_file(document_id, uid, "checkout", str(ws_file))

        # Il file workspace deve essere scrivibile (è in checkout)
        if ws_file.exists():
            self._set_writable(ws_file)

        return arch_file

    def _ext_for_doc_type(self, doc_type: str) -> str:
        from core.file_manager import EXT_FOR_TYPE
        return EXT_FOR_TYPE.get(doc_type, ".SLDPRT")

    # ==================================================================
    #  CONSULTAZIONE (copia senza lock)
    # ==================================================================
    def open_for_consultation(self, document_id: int) -> Path:
        """
        Copia il file dall'archivio alla workspace per sola consultazione.
        Nessun lock, nessun checkout. Registra in workspace_files con role='consultation'.
        """
        doc = self._get_doc_or_raise(document_id)
        dest = self._copy_archive_to_workspace(doc)

        uid = self.current_user["id"]
        self._register_workspace_file(document_id, uid, "consultation", str(dest))

        return dest

    # ==================================================================
    #  WORKSPACE FILES tracking
    # ==================================================================
    def get_workspace_files(self, user_id: int) -> list:
        """Ritorna tutti i file in workspace dell'utente, con info documento."""
        return self.db.fetchall(
            """SELECT wf.*, d.code, d.revision, d.doc_type, d.title, d.state,
                      d.is_locked, d.locked_by, d.checkout_md5,
                      u.full_name as locked_by_name
               FROM workspace_files wf
               JOIN documents d ON d.id = wf.document_id
               LEFT JOIN users u ON u.id = d.locked_by
               WHERE wf.user_id=?
               ORDER BY wf.role, d.code""",
            (user_id,),
        )

    def get_checked_out_by_user(self, user_id: int) -> list:
        """File in checkout dell'utente (retrocompatibilità)."""
        return self.db.fetchall(
            """SELECT d.*, u.full_name as locked_by_name
               FROM documents d
               LEFT JOIN users u ON u.id=d.locked_by
               WHERE d.is_locked=1 AND d.locked_by=?
               ORDER BY d.locked_at DESC""",
            (user_id,),
        )

    def remove_from_workspace(self, document_id: int,
                              delete_file: bool = True) -> bool:
        """Rimuove un file dalla workspace (solo copie senza lock)."""
        uid = self.current_user["id"]
        wf = self.db.fetchone(
            "SELECT * FROM workspace_files WHERE document_id=? AND user_id=?",
            (document_id, uid),
        )
        if not wf:
            return False
        if wf["role"] == "checkout":
            raise PermissionError(
                "Impossibile rimuovere un file in checkout.\n"
                "Usare Check-in o Annulla Checkout."
            )
        if delete_file and wf.get("workspace_path"):
            try:
                p = Path(wf["workspace_path"])
                self._set_writable(p)  # Windows: necessario prima di unlink
                p.unlink()
            except OSError:
                pass
        self._unregister_workspace_file(document_id, uid)
        return True

    # ==================================================================
    #  LOG
    # ==================================================================
    def get_log(self, document_id: int) -> list:
        return self.db.fetchall(
            """SELECT cl.*, u.full_name as user_name
               FROM checkout_log cl
               LEFT JOIN users u ON u.id = cl.user_id
               WHERE cl.document_id=?
               ORDER BY cl.timestamp DESC""",
            (document_id,),
        )

    # ==================================================================
    #  Helpers privati
    # ==================================================================
    def _get_doc_or_raise(self, document_id: int) -> dict:
        doc = self.db.fetchone(
            "SELECT * FROM documents WHERE id=?", (document_id,)
        )
        if not doc:
            raise ValueError("Documento non trovato")
        return doc

    def _check_not_locked(self, doc: dict):
        if doc["is_locked"]:
            locker = self.db.fetchone(
                "SELECT full_name, workstation FROM users WHERE id=?",
                (doc["locked_by"],),
            )
            name = locker["full_name"] if locker else "utente sconosciuto"
            ws = locker.get("workstation", "") if locker else ""
            raise PermissionError(
                f"Documento già in checkout da {name}"
                + (f" ({ws})" if ws else "")
            )

    def _copy_archive_to_workspace(self, doc: dict) -> Path:
        """Copia file dall'archivio alla workspace locale. Ritorna dest path."""
        from config import load_local_config
        from core.file_manager import EXT_FOR_TYPE

        cfg = load_local_config()
        ws = Path(cfg.get("sw_workspace", ""))
        if not ws or not str(ws):
            raise ValueError(
                "Workspace non configurata.\n"
                "Aprire Strumenti → Configurazione SolidWorks e impostare la cartella workspace."
            )
        ws.mkdir(parents=True, exist_ok=True)

        ext = EXT_FOR_TYPE.get(doc.get("doc_type", ""), ".SLDPRT")
        dest = ws / (doc["code"] + ext)

        archive_file = self._archive_file_path(doc)
        if archive_file and archive_file.exists():
            # Se il file esiste già in WS assicuriamoci che sia scrivibile prima di sovrascriverlo
            if dest.exists():
                self._set_writable(dest)
            shutil.copy2(archive_file, dest)
            # Il file in workspace è SEMPRE sola lettura di default;
            # sarà reso scrivibile solo da checkout() per il proprietario del lock.
            self._set_readonly(dest)
        elif doc.get("archive_path"):
            # archive_path registrato ma file fisico mancante
            raise FileNotFoundError(
                f"File archiviato non trovato per {doc['code']} rev.{doc['revision']}:\n"
                f"{archive_file}\n"
                "Verificare che la cartella condivisa sia accessibile."
            )
        else:
            # Documento mai archiviato: non creare file fantasma
            raise PermissionError(
                f"Impossibile fare checkout di {doc['code']} rev.{doc['revision']}:\n"
                "nessun file in archivio.\n"
                "Creare il file in SolidWorks e importarlo in PDM prima del checkout."
            )
        return dest

    def _copy_workspace_to_archive(self, doc: dict, ws_file: Path):
        """Copia file dalla workspace all'archivio (sola lettura)."""
        archive_dir = self.sp.archive_path(doc["code"], doc["revision"])
        archive_dir.mkdir(parents=True, exist_ok=True)

        from core.file_manager import EXT_FOR_TYPE
        ext = EXT_FOR_TYPE.get(doc.get("doc_type", ""), ".SLDPRT")
        dest = archive_dir / (doc["code"] + ext)
        if dest.exists():
            self._set_writable(dest)  # Rimuovi sola lettura prima di sovrascrivere
        shutil.copy2(ws_file, dest)
        self._set_readonly(dest)  # Archivio sempre sola lettura

        rel_path = str(dest.relative_to(self.sp.root))
        uid = self.current_user["id"]
        self.db.execute(
            """UPDATE documents
               SET archive_path=?, file_name=?, file_ext=?,
                   modified_by=?, modified_at=datetime('now')
               WHERE id=?""",
            (rel_path, dest.name, ext, uid, doc["id"]),
        )

    def _checkout_companion_drw(self, doc: dict, uid: int):
        """Mette in checkout il DRW associato al PRT/ASM."""
        if doc["doc_type"] == "Disegno":
            return
        drw = self.db.fetchone(
            "SELECT * FROM documents WHERE code=? AND doc_type='Disegno' "
            "AND revision=?",
            (doc["code"], doc["revision"]),
        )
        if not drw:
            drw = self.db.fetchone(
                "SELECT * FROM documents WHERE code=? AND doc_type='Disegno' "
                "AND state != 'Obsoleto' ORDER BY revision DESC",
                (doc["code"],),
            )
        if drw and drw["state"] not in READONLY_STATES and not drw["is_locked"]:
            try:
                self.checkout(drw["id"], include_drw=False)
            except Exception:
                pass  # DRW checkout non bloccante

    def _register_workspace_file(self, document_id: int, user_id: int,
                                  role: str, workspace_path: str,
                                  parent_checkout_id: int = None):
        """Registra un file nella tabella workspace_files (upsert)."""
        existing = self.db.fetchone(
            "SELECT id FROM workspace_files WHERE document_id=? AND user_id=?",
            (document_id, user_id),
        )
        if existing:
            self.db.execute(
                """UPDATE workspace_files
                   SET role=?, workspace_path=?, copied_at=datetime('now'),
                       parent_checkout_id=?
                   WHERE id=?""",
                (role, workspace_path, parent_checkout_id, existing["id"]),
            )
        else:
            self.db.execute(
                """INSERT INTO workspace_files
                   (document_id, user_id, role, workspace_path, parent_checkout_id)
                   VALUES (?,?,?,?,?)""",
                (document_id, user_id, role, workspace_path, parent_checkout_id),
            )

    def _unregister_workspace_file(self, document_id: int, user_id: int):
        """Rimuove il tracciamento workspace_files."""
        self.db.execute(
            "DELETE FROM workspace_files WHERE document_id=? AND user_id=?",
            (document_id, user_id),
        )

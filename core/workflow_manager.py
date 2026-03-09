# =============================================================================
#  core/workflow_manager.py  –  Gestione workflow documenti  v3.0
# =============================================================================
#  4 stati: In Lavorazione, Rilasciato, In Revisione, Obsoleto
#  Transizioni: In Lavorazione → Rilasciato | In Revisione → Rilasciato
#  «Crea revisione» è un'operazione (non una transizione):
#       Rilasciato → nuova rev in «In Revisione»
#  «Annulla revisione» elimina la rev In Revisione e torna alla precedente
# =============================================================================
from __future__ import annotations
import shutil
import subprocess
import sys
import tempfile
import os
import threading
from pathlib import Path
from typing import Optional, TYPE_CHECKING
if TYPE_CHECKING:
    from core.database import Database

from config import WORKFLOW_TRANSITIONS, WORKFLOW_STATES


class WorkflowManager:
    def __init__(self, db: "Database"):
        self.db = db

    # ------------------------------------------------------------------
    def get_available_transitions(self, current_state: str) -> list[str]:
        return WORKFLOW_TRANSITIONS.get(current_state, [])

    def can_transition(self, from_state: str, to_state: str) -> bool:
        return to_state in WORKFLOW_TRANSITIONS.get(from_state, [])

    # ------------------------------------------------------------------
    def is_latest_revision(self, doc: dict) -> bool:
        """Verifica che il documento sia l'ultima revisione del suo codice+tipo."""
        newer = self.db.fetchone(
            """SELECT id FROM documents
               WHERE code=? AND doc_type=? AND revision>?
               ORDER BY revision DESC LIMIT 1""",
            (doc["code"], doc["doc_type"], doc["revision"]),
        )
        return newer is None

    # ------------------------------------------------------------------
    def change_state(self, document_id: int, new_state: str,
                     user_id: int, notes: str = "",
                     _propagation: bool = False,
                     shared_paths=None) -> bool:
        """
        Cambia stato di un documento con propagazione bidirezionale PRT/ASM ↔ DRW.
        _propagation=True evita loop infinite.
        """
        doc = self.db.fetchone(
            "SELECT * FROM documents WHERE id=?", (document_id,)
        )
        if not doc:
            return False

        current = doc["state"]
        if not self.can_transition(current, new_state):
            raise ValueError(
                f"Transizione non consentita: '{current}' → '{new_state}'"
            )

        # Blocca cambio stato se il documento è in checkout
        if doc.get("is_locked"):
            locker = self.db.fetchone(
                "SELECT full_name FROM users WHERE id=?", (doc["locked_by"],)
            )
            name = locker["full_name"] if locker else f"utente {doc['locked_by']}"
            raise PermissionError(
                f"Impossibile cambiare stato: il documento è in checkout da {name}.\n"
                "Eseguire prima il check-in."
            )

        # Guard: cambio stato solo sull'ultima revisione del codice
        if not _propagation and not self.is_latest_revision(doc):
            raise PermissionError(
                f"Cambio stato consentito solo sull'ultima revisione.\n"
                f"La revisione {doc['revision']} di {doc['code']} non è la più recente."
            )

        # R2: PRT/ASM che passa a 'Rilasciato' deve avere DRW associato
        if (not _propagation
                and doc["doc_type"] in ("Parte", "Assieme")
                and new_state == "Rilasciato"):
            drw_check = self.db.fetchone(
                "SELECT id FROM documents WHERE code=? AND doc_type='Disegno' AND revision=?",
                (doc["code"], doc["revision"]),
            )
            if not drw_check:
                raise PermissionError(
                    f"Impossibile rilasciare {doc['code']} rev.{doc['revision']}:\n"
                    "nessun DRW associato.\n"
                    "Creare prima il disegno (DRW) nella scheda documento."
                )

        # R7: rilascio solo se il file è stato archiviato (checkin eseguito)
        if new_state == "Rilasciato" and not doc.get("archive_path"):
            raise PermissionError(
                f"Impossibile rilasciare {doc['code']} rev.{doc['revision']}:\n"
                "il file non è stato archiviato.\n"
                "Eseguire prima il check-in."
            )
        # R7b: verifica esistenza fisica del file su disco
        if new_state == "Rilasciato" and doc.get("archive_path") and shared_paths:
            phys = shared_paths.root / doc["archive_path"]
            if not phys.exists():
                raise PermissionError(
                    f"Impossibile rilasciare {doc['code']} rev.{doc['revision']}:\n"
                    f"il file archiviato non esiste fisicamente su disco.\n"
                    f"Percorso atteso: {phys}\n"
                    "Verificare che la cartella condivisa sia accessibile e ripetere il check-in."
                )

        # R4: pre-valida che il companion possa seguire la stessa transizione
        if not _propagation:
            _code = doc["code"]
            _rev  = doc["revision"]
            if doc["doc_type"] in ("Parte", "Assieme"):
                _companion = self.db.fetchone(
                    "SELECT * FROM documents WHERE code=? AND doc_type='Disegno' AND revision=?",
                    (_code, _rev),
                )
            elif doc["doc_type"] == "Disegno":
                _companion = self.db.fetchone(
                    "SELECT * FROM documents WHERE code=? "
                    "AND doc_type IN ('Parte','Assieme') AND revision=?",
                    (_code, _rev),
                )
            else:
                _companion = None
            if _companion:
                # R4a: blocca se il companion è in checkout
                if _companion.get("is_locked"):
                    _locker = self.db.fetchone(
                        "SELECT full_name FROM users WHERE id=?",
                        (_companion["locked_by"],),
                    )
                    _lname = _locker["full_name"] if _locker else f"utente {_companion['locked_by']}"
                    raise PermissionError(
                        f"Transizione bloccata: il companion {_companion['doc_type']} "
                        f"({_companion['code']} rev.{_companion['revision']}) "
                        f"è in checkout da {_lname}.\n"
                        "Eseguire prima il check-in del companion."
                    )
                # R4b: blocca se il companion non può seguire la stessa transizione
                if _companion["state"] != new_state:
                    if not self.can_transition(_companion["state"], new_state):
                        raise ValueError(
                            f"Transizione '{current}' \u2192 '{new_state}' bloccata:\n"
                            f"il companion {_companion['doc_type']} è in stato "
                            f"'{_companion['state']}' e non può seguire la stessa transizione.\n"
                            "Allineare prima lo stato del companion."
                        )

        # Aggiorna stato documento
        self.db.execute(
            """UPDATE documents
               SET state=?, modified_by=?, modified_at=datetime('now')
               WHERE id=?""",
            (new_state, user_id, document_id),
        )

        # Storico
        self.db.execute(
            """INSERT INTO workflow_history
               (document_id, from_state, to_state, changed_by, notes)
               VALUES (?,?,?,?,?)""",
            (document_id, current, new_state, user_id, notes),
        )

        # Se rilasciato, rilascia lock e rendi obsoleta la revisione precedente
        if new_state == "Rilasciato":
            self.db.execute(
                """UPDATE documents
                   SET is_locked=0, locked_by=NULL, locked_at=NULL, locked_ws=NULL
                   WHERE id=?""",
                (document_id,),
            )
            self._obsolete_previous_revisions(doc, user_id)

        # Propagazione stato PRT/ASM ↔ DRW (bidirezionale)
        if not _propagation:
            self._propagate_state_to_companion(doc, new_state, user_id, notes,
                                               shared_paths=shared_paths)

        # Genera PDF in background quando un DRW viene rilasciato
        if (new_state == "Rilasciato"
                and doc.get("doc_type") == "Disegno"
                and doc.get("archive_path")
                and shared_paths):
            threading.Thread(
                target=self._generate_pdf_background,
                args=(document_id, doc, shared_paths),
                daemon=True,
            ).start()

        return True

    # ------------------------------------------------------------------
    def _generate_pdf_background(self, document_id: int, doc: dict, shared_paths):
        """Genera PDF del DRW in background via subprocess pdf_worker.py."""
        _WORKER = Path(__file__).parent / "pdf_worker.py"
        _TIMEOUT = 120

        src = shared_paths.root / doc["archive_path"]
        code = doc.get("code", "")
        rev  = doc.get("revision", "")
        pdf_dir = shared_paths.root / "pdf"
        pdf_dir.mkdir(parents=True, exist_ok=True)
        dest = pdf_dir / f"{code}_{rev}.pdf"

        try:
            tmp_fd, tmp_log = tempfile.mkstemp(suffix=".txt", prefix="pdf_")
            os.close(tmp_fd)
            with open(tmp_log, "w") as lf:
                proc = subprocess.Popen(
                    [sys.executable, str(_WORKER), str(src), str(dest)],
                    stdout=lf,
                    stderr=lf,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
            proc.wait(timeout=_TIMEOUT)
            if proc.returncode == 0 and dest.exists():
                try:
                    self.db.set_pdf_path(document_id, str(dest))
                except Exception:
                    pass
        except Exception:
            pass

    # ------------------------------------------------------------------
    def _propagate_state_to_companion(self, doc: dict, new_state: str,
                                       user_id: int, notes: str,
                                       shared_paths=None):
        """Propaga il cambio stato tra PRT/ASM e DRW con stesso codice."""
        code = doc["code"]
        doc_type = doc["doc_type"]

        if doc_type in ("Parte", "Assieme"):
            # Propaga al DRW
            companion = self.db.fetchone(
                "SELECT * FROM documents WHERE code=? AND doc_type='Disegno' AND revision=?",
                (code, doc["revision"]),
            )
        elif doc_type == "Disegno":
            # Propaga al PRT/ASM
            companion = self.db.fetchone(
                "SELECT * FROM documents WHERE code=? AND doc_type IN ('Parte','Assieme') AND revision=?",
                (code, doc["revision"]),
            )
        else:
            return

        if companion and companion["state"] != new_state:
            if self.can_transition(companion["state"], new_state):
                self.change_state(
                    companion["id"], new_state, user_id,
                    notes=f"[Auto] Propagazione da {doc_type} {code}",
                    _propagation=True,
                    shared_paths=shared_paths,
                )

    # ------------------------------------------------------------------
    def _obsolete_previous_revisions(self, doc: dict, user_id: int):
        """Quando una revisione viene rilasciata, le precedenti diventano Obsolete."""
        prev_revs = self.db.fetchall(
            """SELECT id, state, revision FROM documents
               WHERE code=? AND doc_type=? AND id!=? AND state='Rilasciato'""",
            (doc["code"], doc["doc_type"], doc["id"]),
        )
        for prev in prev_revs:
            self.db.execute(
                """UPDATE documents
                   SET state='Obsoleto', modified_by=?, modified_at=datetime('now')
                   WHERE id=?""",
                (user_id, prev["id"]),
            )
            self.db.execute(
                """INSERT INTO workflow_history
                   (document_id, from_state, to_state, changed_by, notes)
                   VALUES (?,?,?,?,?)""",
                (prev["id"], prev["state"], "Obsoleto", user_id,
                 f"Archiviato: nuova revisione {doc['revision']} rilasciata"),
            )

    # ------------------------------------------------------------------
    def get_history(self, document_id: int) -> list:
        return self.db.fetchall(
            """SELECT wh.*, u.full_name as user_name
               FROM workflow_history wh
               LEFT JOIN users u ON u.id = wh.changed_by
               WHERE wh.document_id=?
               ORDER BY wh.changed_at DESC""",
            (document_id,),
        )

    # ------------------------------------------------------------------
    def new_revision(self, document_id: int, user_id: int,
                     new_revision: str, notes: str = "",
                     shared_paths=None) -> int:
        """
        Crea una nuova revisione da un documento Rilasciato:
        - Copia il documento in DB con stato 'In Revisione'
        - Copia il file archiviato dalla vecchia revisione come base
        - Crea anche la nuova revisione del DRW associato (se esiste)
        Ritorna l'ID del nuovo documento.
        """
        doc = self.db.fetchone(
            "SELECT * FROM documents WHERE id=?", (document_id,)
        )
        if not doc:
            raise ValueError("Documento non trovato")

        if doc["state"] != "Rilasciato":
            raise PermissionError(
                "Nuova revisione consentita solo da stato 'Rilasciato'."
            )

        # Crea nuovo documento in DB con stato 'In Revisione'
        new_id = self.db.execute(
            """INSERT INTO documents
               (code, revision, doc_type, title, description,
                state, file_name, file_ext,
                created_by, created_at, modified_by, modified_at,
                machine_id, group_id, doc_level)
               VALUES (?,?,?,?,?,
                       'In Revisione',?,?,
                       ?,datetime('now'),?,datetime('now'),
                       ?,?,?)""",
            (
                doc["code"], new_revision, doc["doc_type"],
                doc["title"], doc["description"],
                doc["file_name"], doc["file_ext"],
                user_id, user_id,
                doc.get("machine_id"), doc.get("group_id"), doc.get("doc_level", 2),
            ),
        )

        # Copia file archiviato come base per la nuova revisione
        if shared_paths and doc.get("archive_path"):
            src = shared_paths.root / doc["archive_path"]
            if src.exists():
                dest_dir = shared_paths.archive_path(doc["code"], new_revision)
                dest_dir.mkdir(parents=True, exist_ok=True)
                dest = dest_dir / src.name
                shutil.copy2(src, dest)
                rel_path = str(dest.relative_to(shared_paths.root))
                self.db.execute(
                    "UPDATE documents SET archive_path=? WHERE id=?",
                    (rel_path, new_id),
                )

        # Storico
        self.db.execute(
            """INSERT INTO workflow_history
               (document_id, from_state, to_state, changed_by, notes)
               VALUES (?,'','In Revisione',?,?)""",
            (new_id, user_id,
             f"Nuova revisione {new_revision} da rev. {doc['revision']}. {notes}"),
        )

        # Crea anche nuova revisione del DRW associato (R6: solo se non esiste già)
        if doc["doc_type"] in ("Parte", "Assieme"):
            drw = self.db.fetchone(
                "SELECT * FROM documents WHERE code=? AND doc_type='Disegno' AND revision=?",
                (doc["code"], doc["revision"]),
            )
            if drw:
                existing_drw_rev = self.db.fetchone(
                    "SELECT id FROM documents WHERE code=? AND doc_type='Disegno' AND revision=?",
                    (doc["code"], new_revision),
                )
                if not existing_drw_rev:
                    self.new_revision(drw["id"], user_id, new_revision,
                                      notes=f"[Auto] Segue nuova rev. {doc['doc_type']}",
                                      shared_paths=shared_paths)

        return new_id

    # ------------------------------------------------------------------
    def cancel_revision(self, document_id: int, user_id: int,
                        shared_paths=None) -> bool:
        """
        Annulla una revisione non ancora rilasciata:
        - Elimina il documento e il relativo file archiviato
        - La revisione precedente resta inalterata
        """
        doc = self.db.fetchone(
            "SELECT * FROM documents WHERE id=?", (document_id,)
        )
        if not doc:
            raise ValueError("Documento non trovato")

        if doc["state"] != "In Revisione":
            raise PermissionError(
                "Annulla revisione consentito solo da stato 'In Revisione'."
            )

        if doc["is_locked"]:
            raise PermissionError(
                "Impossibile annullare: il documento è in checkout."
            )

        # Elimina file archiviato
        if shared_paths and doc.get("archive_path"):
            archive_file = shared_paths.root / doc["archive_path"]
            if archive_file.exists():
                try:
                    archive_file.unlink()
                except OSError:
                    pass

        # Storico prima di eliminare
        self.db.execute(
            """INSERT INTO workflow_history
               (document_id, from_state, to_state, changed_by, notes)
               VALUES (?,?,'ANNULLATO',?,?)""",
            (document_id, doc["state"], user_id,
             f"Revisione {doc['revision']} annullata"),
        )

        # R5: annulla anche il companion della stessa revisione (bidirezionale)
        if doc["doc_type"] in ("Parte", "Assieme"):
            _comp_query = (
                "SELECT * FROM documents WHERE code=? AND doc_type='Disegno' AND revision=?",
                (doc["code"], doc["revision"]),
            )
        elif doc["doc_type"] == "Disegno":
            _comp_query = (
                "SELECT * FROM documents WHERE code=? "
                "AND doc_type IN ('Parte','Assieme') AND revision=?",
                (doc["code"], doc["revision"]),
            )
        else:
            _comp_query = None

        if _comp_query:
            rev_companion = self.db.fetchone(*_comp_query)
            if rev_companion and rev_companion["state"] not in ("Rilasciato", "Obsoleto"):
                if rev_companion.get("is_locked"):
                    _locker = self.db.fetchone(
                        "SELECT full_name FROM users WHERE id=?",
                        (rev_companion["locked_by"],),
                    )
                    _lname = _locker["full_name"] if _locker else f"utente {rev_companion['locked_by']}"
                    raise PermissionError(
                        f"Impossibile annullare la revisione: il companion "
                        f"{rev_companion['doc_type']} ({rev_companion['code']}) "
                        f"è in checkout da {_lname}.\n"
                        "Eseguire prima il check-in del companion."
                    )
                self.cancel_revision(rev_companion["id"], user_id,
                                     shared_paths=shared_paths)

        # Elimina documento
        self.db.execute("DELETE FROM workspace_files WHERE document_id=?", (document_id,))
        self.db.execute("DELETE FROM checkout_log WHERE document_id=?", (document_id,))
        self.db.execute("DELETE FROM documents WHERE id=?", (document_id,))

        return True

    # ------------------------------------------------------------------
    def sync_companion_state(self, document_id: int, target_state: str,
                             user_id: int) -> None:
        """
        Forza lo stato di un documento companion appena creato
        ad allinearsi con il PRT/ASM di riferimento.
        Bypassa i controlli sulle transizioni: usare SOLO alla creazione
        di un companion per garantire coerenza immediata.
        """
        doc = self.db.fetchone("SELECT * FROM documents WHERE id=?", (document_id,))
        if not doc or doc["state"] == target_state:
            return
        current = doc["state"]
        self.db.execute(
            """UPDATE documents
               SET state=?, modified_by=?, modified_at=datetime('now')
               WHERE id=?""",
            (target_state, user_id, document_id),
        )
        self.db.execute(
            """INSERT INTO workflow_history
               (document_id, from_state, to_state, changed_by, notes)
               VALUES (?,?,?,?,?)""",
            (document_id, current, target_state, user_id,
             "[Auto] Allineamento stato alla creazione companion"),
        )

    # ------------------------------------------------------------------
    @staticmethod
    def state_color(state: str) -> str:
        return WORKFLOW_STATES.get(state, {}).get("color", "#9E9E9E")

    @staticmethod
    def all_states() -> list[str]:
        return list(WORKFLOW_STATES.keys())

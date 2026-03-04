# =============================================================================
#  core/workflow_manager.py  –  Gestione workflow documenti  v2.0
# =============================================================================
from __future__ import annotations
import shutil
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
    def change_state(self, document_id: int, new_state: str,
                     user_id: int, notes: str = "",
                     _propagation: bool = False) -> bool:
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
            self._propagate_state_to_companion(doc, new_state, user_id, notes)

        return True

    # ------------------------------------------------------------------
    def _propagate_state_to_companion(self, doc: dict, new_state: str,
                                       user_id: int, notes: str):
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
    def release_document(self, document_id: int, user_id: int,
                         notes: str = "") -> bool:
        """Rilascio diretto (salta la transizione In Revisione)."""
        doc = self.db.fetchone(
            "SELECT * FROM documents WHERE id=?", (document_id,)
        )
        if not doc:
            return False
        if doc["state"] == "Rilasciato":
            return False

        return self.change_state(document_id, "Rilasciato", user_id, notes)

    # ------------------------------------------------------------------
    def new_revision(self, document_id: int, user_id: int,
                     new_revision: str, notes: str = "",
                     shared_paths=None) -> int:
        """
        Crea una nuova revisione:
        - Copia il documento in DB con stato 'In Lavorazione'
        - Copia il file archiviato dalla vecchia revisione come base
        - Crea anche la nuova revisione del DRW associato (se esiste)
        Ritorna l'ID del nuovo documento.
        """
        doc = self.db.fetchone(
            "SELECT * FROM documents WHERE id=?", (document_id,)
        )
        if not doc:
            raise ValueError("Documento non trovato")

        # Crea nuovo documento in DB
        new_id = self.db.execute(
            """INSERT INTO documents
               (code, revision, doc_type, title, description,
                state, file_name, file_ext,
                created_by, created_at, modified_by, modified_at,
                machine_id, group_id, doc_level)
               VALUES (?,?,?,?,?,
                       'In Lavorazione',?,?,
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
               VALUES (?,'','In Lavorazione',?,?)""",
            (new_id, user_id,
             f"Nuova revisione {new_revision} da rev. {doc['revision']}. {notes}"),
        )

        # Crea anche nuova revisione del DRW associato
        if doc["doc_type"] in ("Parte", "Assieme"):
            drw = self.db.fetchone(
                "SELECT * FROM documents WHERE code=? AND doc_type='Disegno' AND revision=?",
                (doc["code"], doc["revision"]),
            )
            if drw:
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

        if doc["state"] in ("Rilasciato", "Obsoleto"):
            raise PermissionError(
                "Impossibile annullare una revisione già rilasciata o obsoleta."
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

        # Elimina documento
        self.db.execute("DELETE FROM workspace_files WHERE document_id=?", (document_id,))
        self.db.execute("DELETE FROM checkout_log WHERE document_id=?", (document_id,))
        self.db.execute("DELETE FROM documents WHERE id=?", (document_id,))

        return True

    # ------------------------------------------------------------------
    @staticmethod
    def state_color(state: str) -> str:
        return WORKFLOW_STATES.get(state, {}).get("color", "#9E9E9E")

    @staticmethod
    def all_states() -> list[str]:
        return list(WORKFLOW_STATES.keys())

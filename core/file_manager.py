# =============================================================================
#  core/file_manager.py  –  Gestione file SolidWorks nell'archivio  v2.0
# =============================================================================
from __future__ import annotations
import shutil
from pathlib import Path
from typing import Optional, TYPE_CHECKING

if TYPE_CHECKING:
    from core.database import Database
    from config import SharedPaths

from config import SW_EXTENSIONS

# Mappa tipo documento → estensione file SW attesa
EXT_FOR_TYPE: dict[str, str] = {
    "Parte":   ".SLDPRT",
    "Assieme": ".SLDASM",
    "Disegno": ".SLDDRW",
}


class FileManager:
    def __init__(self, db: "Database", shared_paths: "SharedPaths",
                 current_user: dict):
        self.db  = db
        self.sp  = shared_paths
        self.cu  = current_user

    # ------------------------------------------------------------------
    # Helpers interni
    # ------------------------------------------------------------------
    def _ext_for_doc(self, doc: dict) -> str:
        """Estensione SW attesa in base al tipo documento."""
        return EXT_FOR_TYPE.get(doc.get("doc_type", ""), ".SLDPRT")

    def _get_workspace(self) -> Optional[Path]:
        """Ritorna la workspace locale configurata (da local_config.json)."""
        from config import load_local_config
        p = load_local_config().get("sw_workspace", "")
        return Path(p) if p else None

    def _require_workspace(self, workspace: Optional[Path] = None) -> Path:
        """Ritorna workspace (passata o configurata); lancia ValueError se assente."""
        ws = workspace or self._get_workspace()
        if not ws:
            raise ValueError(
                "Workspace non configurata.\n"
                "Aprire Strumenti \u2192 Configurazione SolidWorks e impostare la cartella workspace."
            )
        return ws

    # ------------------------------------------------------------------
    # Importazione file nell'archivio (SOLO DA WORKSPACE)
    # ------------------------------------------------------------------
    def import_file(self, source: Path, document_id: int) -> Path:
        """
        Importa un file SolidWorks nell'archivio PDM.
        Il file sorgente DEVE trovarsi nella workspace configurata.
        """
        source = Path(source)
        if not source.exists():
            raise FileNotFoundError(f"File non trovato: {source}")

        # Verifica che il file sia nella workspace
        ws = self._get_workspace()
        if ws:
            try:
                source.relative_to(ws)
            except ValueError:
                raise PermissionError(
                    f"Il file deve trovarsi nella workspace configurata ({ws}).\n"
                    "Non \u00e8 consentito importare file da percorsi esterni."
                )

        doc = self.db.fetchone(
            "SELECT * FROM documents WHERE id=?", (document_id,)
        )
        if not doc:
            raise ValueError("Documento non trovato")

        ext = source.suffix.upper()
        if ext not in SW_EXTENSIONS:
            raise ValueError(f"Estensione non supportata: {ext}")

        # Directory archivio \u2013 rinomina il file con il codice PDM
        arch_dir = self.sp.archive_path(doc["code"], doc["revision"])
        arch_dir.mkdir(parents=True, exist_ok=True)
        dest = arch_dir / (doc["code"] + ext)

        shutil.copy2(source, dest)
        rel_path = str(dest.relative_to(self.sp.root))

        self.db.execute(
            """UPDATE documents
               SET file_name=?, file_ext=?, archive_path=?,
                   modified_by=?, modified_at=datetime('now')
               WHERE id=?""",
            (dest.name, ext, rel_path, self.cu["id"], document_id),
        )
        return dest

    # ------------------------------------------------------------------
    # CREA DA FILE \u2014 unico punto di ingresso da percorso esterno
    # ------------------------------------------------------------------
    def create_from_external_file(self, source: Path, document_id: int,
                                   workspace: Path = None) -> Path:
        """
        Unico entry-point per file provenienti dall'esterno del PDM.
        1. Copia il file nella workspace rinominandolo col codice PDM
        2. NON archivia direttamente \u2014 l'utente far\u00e0 checkin dopo aver lavorato
        Ritorna il path nella workspace.
        """
        ws = self._require_workspace(workspace)
        doc = self.get_document(document_id)
        if not doc:
            raise ValueError("Documento non trovato")

        source = Path(source)
        if not source.exists():
            raise FileNotFoundError(f"File non trovato: {source}")

        ext = source.suffix.upper()
        if ext not in SW_EXTENSIONS:
            raise ValueError(f"Estensione non supportata: {ext}")

        ext_expected = self._ext_for_doc(doc)
        if ext != ext_expected:
            raise ValueError(
                f"Il file selezionato ({ext}) non corrisponde "
                f"al tipo documento '{doc['doc_type']}' (atteso {ext_expected})."
            )

        ws.mkdir(parents=True, exist_ok=True)
        dest = ws / (doc["code"] + ext_expected)
        shutil.copy2(source, dest)
        return dest

    # ------------------------------------------------------------------
    # Utilit\u00e0 DRW
    # ------------------------------------------------------------------
    def find_companion_drw(self, source_file: Path) -> Optional[Path]:
        """Cerca un file .SLDDRW con lo stesso nome base di source_file."""
        for ext in (".SLDDRW", ".slddrw"):
            candidate = source_file.with_suffix(ext)
            if candidate.exists():
                return candidate
        return None

    def get_drw_document(self, prt_asm_doc_id: int) -> Optional[dict]:
        """Trova il documento Disegno che ha stesso codice del PRT/ASM dato."""
        parent = self.get_document(prt_asm_doc_id)
        if not parent:
            return None
        return self.db.fetchone(
            "SELECT * FROM documents WHERE code=? AND doc_type='Disegno'",
            (parent["code"],),
        )

    def get_or_create_drw_document(self, prt_asm_doc_id: int) -> int:
        """Restituisce l'id del Disegno associato; lo crea se non esiste."""
        drw = self.get_drw_document(prt_asm_doc_id)
        if drw:
            return drw["id"]
        parent = self.get_document(prt_asm_doc_id)
        if not parent:
            raise ValueError("Documento padre non trovato")
        return self.create_document(
            parent["code"], parent["revision"], "Disegno",
            parent["title"],
            description=parent.get("description") or "",
            machine_id=parent.get("machine_id"),
            group_id=parent.get("group_id"),
            doc_level=parent.get("doc_level", 2),
            parent_doc_id=prt_asm_doc_id,
        )

    def open_file(self, document_id: int):
        """
        DEPRECATO \u2013 usare open_from_workspace().
        Mantenuto per compatibilit\u00e0: redirige su open_from_workspace.
        """
        self.open_from_workspace(document_id)

    # ------------------------------------------------------------------
    # Operazioni workspace
    # ------------------------------------------------------------------
    def copy_to_workspace(self, source: Path, document_id: int,
                          workspace: Path = None) -> Path:
        """
        Copia *source* nella workspace rinominandolo col codice PDM.
        Verifica coerenza estensione (PRT/ASM/DRW) rispetto al doc_type.
        Ritorna il path destinazione nella workspace.
        """
        ws   = self._require_workspace(workspace)
        doc  = self.get_document(document_id)
        if not doc:
            raise ValueError("Documento non trovato")
        source = Path(source)
        ext_expected = self._ext_for_doc(doc)
        if source.suffix.upper() not in (ext_expected, ext_expected.lower()):
            raise ValueError(
                f"Il file selezionato ({source.suffix.upper()}) non corrisponde "
                f"al tipo documento '{doc['doc_type']}' (atteso {ext_expected})."
            )
        ws.mkdir(parents=True, exist_ok=True)
        dest = ws / (doc["code"] + ext_expected)
        shutil.copy2(source, dest)
        return dest

    def import_from_workspace(self, document_id: int,
                              workspace: Path = None) -> Path:
        """
        Cerca {code}.ext nella workspace e archivia nel PDM.
        Lancia FileNotFoundError se il file non \u00e8 presente nella workspace.
        """
        ws  = self._require_workspace(workspace)
        doc = self.get_document(document_id)
        if not doc:
            raise ValueError("Documento non trovato")
        ext = self._ext_for_doc(doc)
        src = ws / (doc["code"] + ext)
        if not src.exists():
            src_lower = ws / (doc["code"] + ext.lower())
            if src_lower.exists():
                src = src_lower
            else:
                raise FileNotFoundError(
                    f"File non trovato nella workspace:\n{ws / (doc['code'] + ext)}\n\n"
                    "Usare 'Crea in SW' o 'Crea da file' per portare il file in workspace,\n"
                    "poi salvarlo in SolidWorks prima di importare."
                )
        return self.import_file(src, document_id)

    def open_from_workspace(self, document_id: int,
                            workspace: Path = None):
        """
        Prima copia dall'archivio alla workspace (se il file non c'\u00e8 gi\u00e0),
        poi apre con l'applicazione predefinita.
        Non apre MAI direttamente dall'archivio.
        """
        import os
        ws  = self._require_workspace(workspace)
        doc = self.get_document(document_id)
        if not doc:
            raise ValueError("Documento non trovato")
        ext = self._ext_for_doc(doc)
        ws.mkdir(parents=True, exist_ok=True)
        ws_file = ws / (doc["code"] + ext)

        if not ws_file.exists():
            # Retrocompatibilit\u00e0: cerca file con qualsiasi estensione SW/template
            all_exts = (
                list(SW_EXTENSIONS.keys()) +
                [e.lower() for e in SW_EXTENSIONS.keys()] +
                [".prtdot", ".asmdot", ".drwdot",
                 ".PRTDOT", ".ASMDOT", ".DRWDOT",
                 ".sldprt", ".sldasm", ".slddrw"]
            )
            fallback = next(
                (ws / (doc["code"] + x)
                 for x in all_exts
                 if (ws / (doc["code"] + x)).exists()),
                None
            )
            if fallback:
                shutil.move(str(fallback), str(ws_file))
            else:
                if not doc.get("archive_path"):
                    raise ValueError(
                        "Nessun file in archivio e nessun file trovato nella workspace.\n"
                        "Usare 'Crea in SW' o 'Crea da file'."
                    )
                src = self.sp.root / doc["archive_path"]
                if not src.exists():
                    raise FileNotFoundError(f"File archivio non trovato:\n{src}")
                shutil.copy2(src, ws_file)

        os.startfile(str(ws_file))
        return ws_file

    def export_to_workspace(self, document_id: int,
                            workspace: Path = None,
                            include_drw: bool = False) -> list[Path]:
        """
        Copia file dall'ARCHIVIO alla workspace locale.
        Se include_drw, copia anche il DRW associato (se presente in archivio).
        Ritorna la lista dei file copiati nella workspace.
        """
        ws  = self._require_workspace(workspace)
        doc = self.get_document(document_id)
        if not doc:
            raise ValueError("Documento non trovato")

        results: list[Path] = []
        ws.mkdir(parents=True, exist_ok=True)

        # File principale
        if doc.get("archive_path"):
            src = self.sp.root / doc["archive_path"]
            if src.exists():
                ext  = self._ext_for_doc(doc)
                dest = ws / (doc["code"] + ext)
                shutil.copy2(src, dest)
                results.append(dest)

        # DRW associato
        if include_drw and doc["doc_type"] in ("Parte", "Assieme"):
            drw_doc = self.get_drw_document(document_id)
            if drw_doc and drw_doc.get("archive_path"):
                drw_src = self.sp.root / drw_doc["archive_path"]
                if drw_src.exists():
                    drw_dest = ws / (drw_doc["code"] + ".SLDDRW")
                    shutil.copy2(drw_src, drw_dest)
                    results.append(drw_dest)

        return results

    def export_from_workspace(self, document_id: int, dest_folder: Path,
                              workspace: Path = None) -> list[Path]:
        """
        Esporta file dalla WORKSPACE verso un percorso esterno.
        Non accede direttamente all'archivio \u2014 il file deve essere prima
        in workspace (via consultazione o checkout).
        """
        ws = self._require_workspace(workspace)
        doc = self.get_document(document_id)
        if not doc:
            raise ValueError("Documento non trovato")

        results: list[Path] = []
        ext = self._ext_for_doc(doc)
        ws_file = ws / (doc["code"] + ext)

        if ws_file.exists():
            dest_folder = Path(dest_folder)
            dest_folder.mkdir(parents=True, exist_ok=True)
            dest = dest_folder / ws_file.name
            shutil.copy2(ws_file, dest)
            results.append(dest)
        else:
            raise FileNotFoundError(
                f"File non trovato nella workspace:\n{ws_file}\n\n"
                "Scaricare prima il file in workspace tramite Consultazione o Checkout."
            )
        return results

    # ------------------------------------------------------------------
    # Ricerca file nella workspace dell'utente
    # ------------------------------------------------------------------
    def scan_workspace(self, user_workspace: Path) -> list[Path]:
        """Ritorna la lista dei file SW nella workspace dell'utente."""
        result = []
        for ext in SW_EXTENSIONS:
            result.extend(user_workspace.rglob(f"*{ext}"))
            result.extend(user_workspace.rglob(f"*{ext.lower()}"))
        return sorted(set(result))

    # ------------------------------------------------------------------
    # Documento
    # ------------------------------------------------------------------
    def create_document(self, code: str, revision: str, doc_type: str,
                        title: str, description: str = "",
                        machine_id: int = None, group_id: int = None,
                        doc_level: int = 2, parent_doc_id: int = None) -> int:
        return self.db.execute(
            """INSERT INTO documents
               (code, revision, doc_type, title, description, state, created_by,
                machine_id, group_id, doc_level, parent_doc_id)
               VALUES (?,?,?,?,?,'In Lavorazione',?,?,?,?,?)""",
            (code, revision, doc_type, title, description, self.cu["id"],
             machine_id, group_id, doc_level, parent_doc_id),
        )

    def create_from_template(self, document_id: int,
                              workspace_folder: Path,
                              also_drw: bool = False
                              ) -> "tuple[Path, Optional[Path]]":
        """
        Copia il template SW nella workspace rinominandolo col codice PDM.
        Se also_drw=True e il documento \u00e8 PRT/ASM, crea anche il file DRW
        (e il relativo documento in DB se assente).
        Ritorna (path_principale, path_drw_o_None).
        """
        from config import load_local_config
        cfg = load_local_config()
        doc = self.get_document(document_id)
        if not doc:
            raise ValueError("Documento non trovato")

        key_map = {
            "Parte":   "sw_template_prt",
            "Assieme": "sw_template_asm",
            "Disegno": "sw_template_drw",
        }
        tpl_key  = key_map.get(doc["doc_type"], "")
        tpl_path = Path(cfg.get(tpl_key, "") or "")
        if not tpl_path.exists():
            raise FileNotFoundError(
                f"Template non configurato per '{doc['doc_type']}'.\n"
                "Configurarlo in Strumenti \u2192 Configurazione SolidWorks."
            )

        sw_ext = EXT_FOR_TYPE.get(doc["doc_type"], ".SLDPRT")
        dest = Path(workspace_folder) / (doc["code"] + sw_ext)
        shutil.copy2(tpl_path, dest)

        drw_dest: Optional[Path] = None
        if also_drw and doc["doc_type"] in ("Parte", "Assieme"):
            drw_tpl_path = Path(cfg.get("sw_template_drw", "") or "")
            if drw_tpl_path.exists():
                drw_dest = Path(workspace_folder) / (doc["code"] + ".SLDDRW")
                shutil.copy2(drw_tpl_path, drw_dest)
                self.get_or_create_drw_document(document_id)

        return dest, drw_dest

    def update_document(self, document_id: int, title: str,
                        description: str) -> bool:
        self.db.execute(
            """UPDATE documents
               SET title=?, description=?, modified_by=?, modified_at=datetime('now')
               WHERE id=?""",
            (title, description, self.cu["id"], document_id),
        )
        return True

    def delete_document(self, document_id: int) -> bool:
        """Elimina un documento (solo se non rilasciato e non in checkout)."""
        doc = self.db.fetchone(
            "SELECT * FROM documents WHERE id=?", (document_id,)
        )
        if not doc:
            return False
        if doc["state"] == "Rilasciato":
            raise PermissionError("Impossibile eliminare un documento Rilasciato")
        if doc["is_locked"]:
            raise PermissionError("Impossibile eliminare un documento in checkout")

        self.db.execute("DELETE FROM documents WHERE id=?", (document_id,))
        return True

    def get_document(self, document_id: int) -> Optional[dict]:
        return self.db.fetchone(
            """SELECT d.*, uc.full_name as created_by_name,
                      um.full_name as modified_by_name,
                      ul.full_name as locked_by_name
               FROM documents d
               LEFT JOIN users uc ON uc.id=d.created_by
               LEFT JOIN users um ON um.id=d.modified_by
               LEFT JOIN users ul ON ul.id=d.locked_by
               WHERE d.id=?""",
            (document_id,),
        )

    def search_documents(self, text: str = "", doc_type: str = "",
                         state: str = "", code: str = "") -> list:
        conditions, params = ["1=1"], []
        if text:
            conditions.append(
                "(d.title LIKE ? OR d.description LIKE ? OR d.code LIKE ?)"
            )
            params += [f"%{text}%", f"%{text}%", f"%{text}%"]
        if doc_type:
            conditions.append("d.doc_type=?"); params.append(doc_type)
        if state:
            conditions.append("d.state=?"); params.append(state)
        if code:
            conditions.append("d.code LIKE ?"); params.append(f"%{code}%")

        sql = f"""
            SELECT d.*,
                   uc.full_name as created_by_name,
                   ul.full_name as locked_by_name
            FROM documents d
            LEFT JOIN users uc ON uc.id=d.created_by
            LEFT JOIN users ul ON ul.id=d.locked_by
            WHERE {' AND '.join(conditions)}
            ORDER BY d.code, d.revision
        """
        return self.db.fetchall(sql, params)

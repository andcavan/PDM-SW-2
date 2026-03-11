# =============================================================================
#  core/asm_manager.py  –  Gestione struttura assieme (BOM)
# =============================================================================
from __future__ import annotations
from pathlib import Path
from typing import Optional, TYPE_CHECKING

if TYPE_CHECKING:
    from core.database import Database


class AsmManager:
    def __init__(self, db: "Database"):
        self.db = db

    # ------------------------------------------------------------------
    # Aggiunta / rimozione componenti
    # ------------------------------------------------------------------
    def add_component(self, parent_id: int, child_id: int,
                      quantity: float = 1.0, position: str = "",
                      notes: str = "") -> int:
        if parent_id == child_id:
            raise ValueError("Un documento non può contenere se stesso")

        existing = self.db.fetchone(
            "SELECT id FROM asm_components WHERE parent_id=? AND child_id=?",
            (parent_id, child_id),
        )
        if existing:
            self.db.execute(
                "UPDATE asm_components SET quantity=?, position=?, notes=? WHERE id=?",
                (quantity, position, notes, existing["id"]),
            )
            return existing["id"]

        return self.db.execute(
            """INSERT INTO asm_components (parent_id, child_id, quantity, position, notes)
               VALUES (?,?,?,?,?)""",
            (parent_id, child_id, quantity, position, notes),
        )

    def remove_component(self, parent_id: int, child_id: int):
        self.db.execute(
            "DELETE FROM asm_components WHERE parent_id=? AND child_id=?",
            (parent_id, child_id),
        )

    def update_quantity(self, parent_id: int, child_id: int, quantity: float):
        self.db.execute(
            "UPDATE asm_components SET quantity=? WHERE parent_id=? AND child_id=?",
            (quantity, parent_id, child_id),
        )

    # ------------------------------------------------------------------
    # Lettura BOM
    # ------------------------------------------------------------------
    def get_components(self, parent_id: int) -> list:
        """Componenti diretti (primo livello) dell'assieme."""
        return self.db.fetchall(
            """SELECT ac.*, d.code, d.revision, d.doc_type, d.title, d.state
               FROM asm_components ac
               JOIN documents d ON d.id = ac.child_id
               WHERE ac.parent_id=?
               ORDER BY d.code""",
            (parent_id,),
        )

    def get_bom_flat(self, parent_id: int, _visited: Optional[set] = None) -> list:
        """BOM appiattita (tutti i livelli)."""
        if _visited is None:
            _visited = set()
        if parent_id in _visited:
            return []
        _visited.add(parent_id)

        result = []
        for comp in self.get_components(parent_id):
            result.append(comp)
            if comp["doc_type"] == "Assieme":
                children = self.get_bom_flat(comp["child_id"], _visited)
                result.extend(children)
        return result

    def get_where_used(self, child_id: int) -> list:
        """In quali assiemi è usato questo documento."""
        return self.db.fetchall(
            """SELECT ac.*, d.code, d.revision, d.title, d.state
               FROM asm_components ac
               JOIN documents d ON d.id = ac.parent_id
               WHERE ac.child_id=?
               ORDER BY d.code""",
            (child_id,),
        )

    # ------------------------------------------------------------------
    # Import da SolidWorks (COM API)
    # ------------------------------------------------------------------
    def import_from_sw_asm(self, asm_file: Path,
                            parent_doc_id: int) -> int:
        """
        Legge la struttura dell'assieme via COM API e popola asm_components.
        Usa GetComponents(True) per il primo livello; ricorre nei sotto-assiemi
        aprendo ciascun file .SLDASM trovato.
        Ritorna il totale di relazioni inserite/aggiornate.
        """
        try:
            import win32com.client as win32
            import pythoncom
        except ImportError:
            raise ImportError("pywin32 non installato.\nEseguire: pip install pywin32")

        pythoncom.CoInitialize()
        sw = win32.Dispatch("SldWorks.Application")
        sw.Visible = False

        try:
            count = self._import_asm_recursive(
                sw, pythoncom, Path(asm_file), parent_doc_id, depth=0, opened=set()
            )
        finally:
            pythoncom.CoUninitialize()
        return count

    def import_bom_from_active_doc(self, parent_doc_id: int,
                                    expected_path: "Path | None" = None) -> int:
        """
        Legge la struttura BOM dal documento ATTIVO in SolidWorks.
        Usa GetActiveObject (non Dispatch) per collegarsi all'istanza aperta.
        Cancella le relazioni vecchie e le ricrea.
        Ritorna il numero di relazioni create.

        Se expected_path è fornito, verifica che il documento attivo corrisponda
        al file atteso prima di procedere (sicurezza coerenza PDM↔SW).
        """
        try:
            import win32com.client as win32
        except ImportError:
            raise ImportError("pywin32 non installato")

        try:
            sw = win32.GetActiveObject("SldWorks.Application")
        except Exception:
            raise RuntimeError(
                "SolidWorks non è in esecuzione.\n"
                "Aprire il file in SolidWorks prima di importare la struttura."
            )

        model = sw.ActiveDoc
        if model is None:
            raise RuntimeError("Nessun documento attivo in SolidWorks.")

        # Valida che il documento attivo corrisponda al file atteso
        if expected_path is not None:
            try:
                active_path = model.GetPathName
                if callable(active_path):
                    active_path = active_path()
                if active_path and Path(str(active_path)).resolve() != Path(expected_path).resolve():
                    raise RuntimeError(
                        f"Il documento attivo in SolidWorks ({Path(str(active_path)).name}) "
                        f"non corrisponde al file da archiviare ({Path(expected_path).name}).\n\n"
                        "Assicurarsi che il file corretto sia attivo in SolidWorks prima del check-in."
                    )
            except RuntimeError:
                raise
            except Exception:
                pass

        # Leggi componenti primo livello
        # GetComponents è una proprietà con GetActiveObject
        try:
            raw = model.GetComponents(True)
        except TypeError:
            raw = model.GetComponents
        if raw is None:
            # Cancella le vecchie relazioni
            self.db.execute(
                "DELETE FROM asm_components WHERE parent_id=?",
                (parent_doc_id,),
            )
            return 0

        components = list(raw) if not isinstance(raw, (list, tuple)) else raw

        # Raggruppa per nome file → quantità
        from collections import Counter
        path_counter: Counter = Counter()
        path_map: dict = {}

        for comp in components:
            fp = ""
            try:
                fp = comp.GetPathName
                if callable(fp):
                    fp = fp()
                fp = str(fp or "").strip()
            except Exception:
                try:
                    fp = str(comp.GetPathName or "").strip()
                except Exception:
                    fp = ""
            if not fp:
                continue
            k = fp.upper()
            path_counter[k] += 1
            path_map[k] = fp

        # Cancella vecchie relazioni e ricrea
        self.db.execute(
            "DELETE FROM asm_components WHERE parent_id=?",
            (parent_doc_id,),
        )

        count = 0
        for k, qty in path_counter.items():
            fp = path_map[k]
            code = Path(fp).stem

            # Cerca il documento per codice (qualsiasi tipo tranne Disegno)
            doc = self.db.fetchone(
                "SELECT id, doc_type FROM documents "
                "WHERE code=? AND doc_type != 'Disegno' "
                "AND state != 'Obsoleto' "
                "ORDER BY revision DESC",
                (code,),
            )
            if not doc:
                continue

            self.add_component(parent_doc_id, doc["id"], float(qty))
            count += 1

        return count

    def _open_sw_doc(self, sw, pythoncom, file_path: Path):
        """Apre un documento SW e ritorna il model object."""
        import win32com.client as win32
        doc_type = 2 if file_path.suffix.upper() == ".SLDASM" else 1
        errors_v   = win32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        warnings_v = win32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        model = sw.OpenDoc6(str(file_path), doc_type, 1, "", errors_v, warnings_v)
        return model

    def _import_asm_recursive(self, sw, pythoncom, asm_file: Path,
                               parent_doc_id: int, depth: int,
                               opened: set) -> int:
        """
        Apre asm_file, chiama GetComponents(True) per il primo livello,
        aggiunge le relazioni al padre, poi ricorre nei sotto-assiemi.
        """
        if depth > 15:
            return 0
        key = str(asm_file).upper()
        if key in opened:
            return 0
        opened.add(key)

        model = None
        count = 0
        try:
            model = self._open_sw_doc(sw, pythoncom, asm_file)
            if not model:
                return 0

            # GetComponents(True) → lista componenti primo livello
            raw = model.GetComponents(True)
            if raw is None:
                return 0
            components = list(raw) if not isinstance(raw, (list, tuple)) else raw

            # Raggruppa per path per calcolare quantità
            from collections import Counter
            path_counter: Counter = Counter()
            path_map: dict = {}
            for comp in components:
                try:
                    fp = str(comp.GetPathName() or "").strip()
                except Exception:
                    try:
                        fp = str(comp.GetPathName or "").strip()
                    except Exception:
                        fp = ""
                if not fp:
                    continue
                k = fp.upper()
                path_counter[k] += 1
                path_map[k] = (comp, fp)

            for k, qty in path_counter.items():
                comp, fp = path_map[k]
                code = Path(fp).stem

                doc = self.db.fetchone(
                    "SELECT id, doc_type FROM documents WHERE code=?", (code,)
                )
                if not doc:
                    continue

                self.add_component(parent_doc_id, doc["id"], float(qty))
                count += 1

                # Ricorre nei sotto-assiemi
                if doc["doc_type"] == "Assieme" and Path(fp).exists():
                    count += self._import_asm_recursive(
                        sw, pythoncom, Path(fp), doc["id"], depth + 1, opened
                    )
        finally:
            if model:
                try:
                    title = model.GetTitle()
                    if title:
                        sw.CloseDoc(title)
                except Exception:
                    pass
        return count

    # ------------------------------------------------------------------
    # Lettura struttura ASM senza modifiche al DB (per wizard importazione)
    # ------------------------------------------------------------------
    def read_asm_tree_from_active(self) -> list:
        """
        Legge la struttura ASM dal documento ATTIVO in SolidWorks.
        Usa GetComponents(False) sulla radice per ottenere TUTTI i componenti
        a tutti i livelli in una sola chiamata COM, senza ricorsione su GetModelDoc2.
        NON scrive nulla nel DB.
        Ritorna lista piatta (visita DFS) di dict:
          {name, path, type:'PRT'|'ASM', depth, parent_path, quantity}
        Il primo elemento è sempre il documento radice (l'ASM stesso).
        """
        import logging
        from collections import Counter, defaultdict

        try:
            import win32com.client as win32
        except ImportError:
            raise ImportError("pywin32 non installato.\nEseguire: pip install pywin32")

        try:
            sw = win32.GetActiveObject("SldWorks.Application")
        except Exception:
            raise RuntimeError(
                "SolidWorks non è in esecuzione.\n"
                "Aprire il file assieme in SolidWorks prima di avviare il wizard."
            )

        model = sw.ActiveDoc
        if model is None:
            raise RuntimeError("Nessun documento attivo in SolidWorks.")

        asm_path = ""
        try:
            v = model.GetPathName
            if callable(v):
                v = v()
            asm_path = str(v or "").strip()
        except Exception:
            pass

        result = []

        # Radice = l'assieme stesso
        asm_name = Path(asm_path).stem if asm_path else "Assieme radice"
        result.append({
            "name":        asm_name,
            "path":        asm_path,
            "type":        "ASM",
            "depth":       0,
            "parent_path": None,
            "quantity":    1,
        })

        # GetComponents(False) → TUTTI i componenti a tutti i livelli in un colpo solo.
        # SolidWorks risolve internamente la gerarchia; nessuna ricorsione su GetModelDoc2.
        try:
            raw = model.GetComponents(False)
        except Exception as e:
            logging.warning("GetComponents(False) fallito: %s", e)
            return result

        if raw is None:
            return result

        try:
            all_comps = list(raw)
        except Exception:
            return result

        if not all_comps:
            return result

        def _get_path(obj) -> str:
            try:
                v = obj.GetPathName
                if callable(v):
                    v = v()
                return str(v or "").strip()
            except Exception:
                return ""

        # Prima passata: conta occorrenze (parent_key, child_key) e mantieni ordine DFS.
        # GetComponents(False) restituisce i componenti in ordine DFS come SolidWorks li visita.
        pair_count: Counter  = Counter()
        pair_first: dict     = {}          # (pk, ck) → (fp, pp)
        children_map: dict   = defaultdict(list)   # pk → [(pk, ck)] in ordine
        pair_order_seen: set = set()

        for comp in all_comps:
            fp = _get_path(comp)
            if not fp:
                continue

            # Padre del componente via GetParent(); None = figlio diretto della radice
            pp = asm_path
            try:
                parent_comp = comp.GetParent()
                if parent_comp is not None:
                    pp_c = _get_path(parent_comp)
                    if pp_c:
                        pp = pp_c
            except Exception:
                pass

            pk = pp.upper()
            ck = fp.upper()

            pair_count[(pk, ck)] += 1

            if (pk, ck) not in pair_order_seen:
                pair_order_seen.add((pk, ck))
                pair_first[(pk, ck)] = (fp, pp)
                children_map[pk].append((pk, ck))

        # Visita DFS per costruire la lista piatta con depth e quantità per-istanza corrette.
        # parent_total = quante istanze del padre esistono nel documento → normalizza qty figli.
        def dfs(parent_path: str, depth: int, parent_total: int):
            pk = parent_path.upper()
            for (ppk, cck) in children_map.get(pk, []):
                if (ppk, cck) not in pair_first:
                    continue
                fp, pp = pair_first[(ppk, cck)]
                child_total = pair_count[(ppk, cck)]
                per_instance = max(1, child_total // parent_total) if parent_total > 0 else 1
                suffix = Path(fp).suffix.upper()
                node_type = "ASM" if suffix == ".SLDASM" else "PRT"
                result.append({
                    "name":        Path(fp).stem,
                    "path":        fp,
                    "type":        node_type,
                    "depth":       depth,
                    "parent_path": pp,
                    "quantity":    per_instance,
                })
                if node_type == "ASM":
                    dfs(fp, depth + 1, child_total)

        dfs(asm_path, 1, 1)
        return result

    def _process_sw_component(self, *args, **kwargs) -> int:
        """Stub mantenuto per compatibilità — non più usato."""
        return 0

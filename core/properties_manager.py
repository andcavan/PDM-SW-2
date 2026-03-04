# =============================================================================
#  core/properties_manager.py  –  Import/export proprietà SolidWorks
# =============================================================================
from __future__ import annotations
from pathlib import Path
from typing import Optional, TYPE_CHECKING

if TYPE_CHECKING:
    from core.database import Database


# Proprietà standard SolidWorks
SW_STANDARD_PROPS = [
    "Description", "PartNo", "DrawnBy", "DrawnDate",
    "CheckedBy", "CheckedDate", "EngineeringApproval",
    "ManufacturingApproval", "QualityApproval",
    "Material", "Finish", "Weight", "Density",
    "Revision", "Title", "Project", "Company",
]


class PropertiesManager:
    def __init__(self, db: "Database"):
        self.db = db

    # ------------------------------------------------------------------
    # DB
    # ------------------------------------------------------------------
    def save_properties(self, document_id: int, props: dict):
        """Salva un dizionario {nome: valore} nel DB."""
        for name, value in props.items():
            existing = self.db.fetchone(
                "SELECT id FROM document_properties WHERE document_id=? AND prop_name=?",
                (document_id, name),
            )
            if existing:
                self.db.execute(
                    "UPDATE document_properties SET prop_value=? WHERE id=?",
                    (str(value), existing["id"]),
                )
            else:
                self.db.execute(
                    """INSERT INTO document_properties (document_id, prop_name, prop_value)
                       VALUES (?,?,?)""",
                    (document_id, name, str(value)),
                )

    def get_properties(self, document_id: int) -> dict:
        rows = self.db.fetchall(
            "SELECT prop_name, prop_value FROM document_properties WHERE document_id=?",
            (document_id,),
        )
        return {r["prop_name"]: r["prop_value"] for r in rows}

    def delete_property(self, document_id: int, prop_name: str):
        self.db.execute(
            "DELETE FROM document_properties WHERE document_id=? AND prop_name=?",
            (document_id, prop_name),
        )

    # ------------------------------------------------------------------
    # COM – SolidWorks API (richiede pywin32 e SW installato)
    # ------------------------------------------------------------------
    def read_from_sw_file(self, file_path: Path, file_name: str = None) -> dict:
        """
        Legge le proprietà da un file SolidWorks.
        Se SW è già in esecuzione e il file è aperto, usa il doc esistente.
        Cerca per: 1) path esatto, 2) file_name tra i doc aperti, 3) OpenDoc6.
        """
        try:
            import win32com.client as win32
        except ImportError:
            raise ImportError(
                "pywin32 non installato. "
                "Eseguire: pip install pywin32"
            )

        props = {}
        opened_by_us = False
        sw = None
        model = None

        try:
            # Prova a connettersi all'istanza SW già in esecuzione
            try:
                sw = win32.GetActiveObject("SldWorks.Application")
            except Exception:
                sw = win32.Dispatch("SldWorks.Application")

            if sw is None:
                props["_error"] = "Impossibile connettersi a SolidWorks"
                return props

            search_name = (file_name or file_path.name).upper()

            # 1) Prova il documento attivo (caso più comune)
            #    L'utente ha premuto "Importa da SW" → vuole le props
            #    del doc attivo in SolidWorks, indipendentemente dal nome.
            try:
                active = sw.ActiveDoc
                if active is not None:
                    model = active
            except Exception:
                pass

            # 2) Cerca per path esatto
            if model is None:
                file_str = str(file_path)
                try:
                    model = sw.GetOpenDocumentByName(file_str)
                except Exception:
                    pass

            # 3) Cerca per nome file tra i documenti aperti
            #    Itera i doc, trova il path completo, poi usa
            #    GetOpenDocumentByName per ottenere un dispatch COM corretto
            if model is None:
                found_path = None
                try:
                    doc_iter = sw.GetFirstDocument()
                    while doc_iter is not None:
                        try:
                            fp = doc_iter.GetPathName
                            if isinstance(fp, str) and fp:
                                if Path(fp).name.upper() == search_name:
                                    found_path = fp
                                    break
                        except Exception:
                            pass
                        try:
                            doc_iter = doc_iter.GetNext
                            if doc_iter is None or not hasattr(doc_iter, 'GetPathName'):
                                break
                        except Exception:
                            break
                except Exception:
                    pass

                if found_path:
                    try:
                        model = sw.GetOpenDocumentByName(found_path)
                    except Exception:
                        pass

            # 4) Prova ad aprire il file dal path (fallback)
            if model is None:
                ext = file_path.suffix.upper()
                doc_type_map = {
                    ".SLDPRT": 1,   # swDocPART
                    ".SLDASM": 2,   # swDocASSEMBLY
                    ".SLDDRW": 3,   # swDocDRAWING
                }
                doc_type = doc_type_map.get(ext, 1)
                arg_err = win32.VARIANT(win32.pythoncom.VT_BYREF | win32.pythoncom.VT_I4, 0)
                arg_warn = win32.VARIANT(win32.pythoncom.VT_BYREF | win32.pythoncom.VT_I4, 0)
                model = sw.OpenDoc6(
                    str(file_path), doc_type,
                    1,   # swOpenDocOptions_Silent
                    "", arg_err, arg_warn
                )
                if model:
                    opened_by_us = True

            if model is None:
                props["_error"] = (
                    f"File '{search_name}' non trovato in SolidWorks.\n"
                    "Aprire il file in SolidWorks prima di importare."
                )
                return props

            # Leggi proprietà custom (configurazione default "")
            props = self._read_custom_props(model)

        except Exception as e:
            props["_error"] = str(e)
        finally:
            if opened_by_us and model and sw:
                try:
                    sw.CloseDoc(model.GetTitle())
                except Exception:
                    pass
        return props

    @staticmethod
    def _read_custom_props(model) -> dict:
        """
        Legge le proprietà custom dal modello SW.
        Legge sia dalla config "" (file-level) che dalle configurazioni specifiche.
        Nota: con GetActiveObject, i metodi COM possono essere esposti come
        proprietà (senza parentesi): GetNames → proprietà, Get → proprietà.
        """
        props = {}

        ext = model.Extension

        # Raccogli lista configurazioni da leggere: "" (file-level) + tutte le config
        config_names = [""]
        try:
            cfgs = model.GetConfigurationNames
            if cfgs:
                if isinstance(cfgs, str):
                    config_names.append(cfgs)
                else:
                    config_names.extend(cfgs)
        except Exception:
            pass

        for cfg_name in config_names:
            try:
                mgr = ext.CustomPropertyManager(cfg_name)
                if isinstance(mgr, tuple):
                    mgr = mgr[0] if mgr else None
                if mgr is None:
                    continue
            except Exception:
                continue

            try:
                count = mgr.Count
                if not count or count <= 0:
                    continue
            except Exception:
                continue

            # GetNames è una proprietà (non metodo) con GetActiveObject
            names = None
            try:
                names = mgr.GetNames
            except Exception:
                pass
            if names is None:
                try:
                    names = mgr.GetNames()
                except Exception:
                    pass
            if not names:
                continue

            # Se è una stringa singola, wrappala in lista
            if isinstance(names, str):
                names = [names]

            for name in names:
                if name in props:
                    continue  # file-level ha priorità
                try:
                    # Get è una proprietà (non metodo) con GetActiveObject
                    val = mgr.Get(name)
                    props[name] = str(val) if val else ""
                except TypeError:
                    # Se Get non è callable come proprietà, prova come metodo
                    try:
                        val = mgr.Get
                        props[name] = str(val) if val else ""
                    except Exception:
                        props[name] = ""
                except Exception:
                    props[name] = ""

        return props

    def write_to_sw_file(self, file_path: Path, props: dict):
        """Scrive le proprietà in un file SolidWorks via COM."""
        try:
            import win32com.client as win32
        except ImportError:
            raise ImportError("pywin32 non installato")

        opened_by_us = False
        sw = None
        model = None

        try:
            # Prova a connettersi all'istanza SW già in esecuzione
            try:
                sw = win32.GetActiveObject("SldWorks.Application")
            except Exception:
                sw = win32.Dispatch("SldWorks.Application")

            if sw is None:
                raise RuntimeError("Impossibile connettersi a SolidWorks")

            # Cerca se il file è già aperto
            file_str = str(file_path)
            model = sw.GetOpenDocumentByName(file_str)

            if model is None:
                ext = file_path.suffix.upper()
                doc_type_map = {
                    ".SLDPRT": 1, ".SLDASM": 2, ".SLDDRW": 3
                }
                doc_type = doc_type_map.get(ext, 1)
                arg_err = win32.VARIANT(win32.pythoncom.VT_BYREF | win32.pythoncom.VT_I4, 0)
                arg_warn = win32.VARIANT(win32.pythoncom.VT_BYREF | win32.pythoncom.VT_I4, 0)
                model = sw.OpenDoc6(
                    file_str, doc_type,
                    0, "", arg_err, arg_warn
                )
                if model:
                    opened_by_us = True

            if model is None:
                raise RuntimeError(f"Impossibile aprire: {file_path.name}")

            mgr = model.Extension.CustomPropertyManager("")
            for name, value in props.items():
                # swCustomInfoText=30, swCustomPropertyReplaceValue=2
                mgr.Add3(name, 30, str(value), 2)

            model.Save()

        finally:
            if opened_by_us and model and sw:
                try:
                    sw.CloseDoc(model.GetTitle())
                except Exception:
                    pass

    # ------------------------------------------------------------------
    # Export Excel
    # ------------------------------------------------------------------
    def export_to_excel(self, document_id: int, dest: Path):
        try:
            import openpyxl
        except ImportError:
            raise ImportError("openpyxl non installato")

        props = self.get_properties(document_id)
        doc   = self.db.fetchone(
            "SELECT code, revision, title FROM documents WHERE id=?",
            (document_id,),
        )
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Proprietà"
        ws.append(["Codice", doc["code"] if doc else ""])
        ws.append(["Revisione", doc["revision"] if doc else ""])
        ws.append(["Titolo", doc["title"] if doc else ""])
        ws.append([])
        ws.append(["Proprietà", "Valore"])
        for k, v in props.items():
            ws.append([k, v])
        wb.save(str(dest))

    def import_from_excel(self, document_id: int, src: Path) -> int:
        try:
            import openpyxl
        except ImportError:
            raise ImportError("openpyxl non installato")

        wb = openpyxl.load_workbook(str(src))
        ws = wb.active
        props = {}
        header_found = False
        for row in ws.iter_rows(values_only=True):
            if not header_found:
                if row and str(row[0]).lower() == "proprietà":
                    header_found = True
                continue
            if row and row[0]:
                props[str(row[0])] = str(row[1]) if row[1] is not None else ""
        self.save_properties(document_id, props)
        return len(props)

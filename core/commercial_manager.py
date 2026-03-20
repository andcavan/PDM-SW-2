# =============================================================================
#  core/commercial_manager.py  –  Gestione articoli commerciali/normalizzati
# =============================================================================
from __future__ import annotations

import hashlib
import os
import shutil
import socket
import stat as stat_mod
from datetime import datetime
from pathlib import Path
from typing import Optional, TYPE_CHECKING

from config import COMMERCIAL_ITEM_TYPES, COMMERCIAL_WORKFLOW_TRANSITIONS, COMMERCIAL_READONLY_STATES
from core.commercial_coding_config import CommercialCodingConfig

if TYPE_CHECKING:
    from core.database import Database
    from config import SharedPaths


class CommercialManager:
    """
    Logica di business per articoli commerciali e normalizzati.

    Gestisce:
      - Categorie e sottocategorie
      - Codifica automatica ({L}-{CAT}-{SUB}-{NUM:4})
      - CRUD articoli commerciali
      - Fornitori e associazione articolo-fornitore
      - Checkout / Checkin file SolidWorks (opzionale)
      - Proprietà SolidWorks (import/export via PropertiesManager)
      - Integrazione BOM CAD (asm_commercial_components)
    """

    def __init__(self, db: "Database"):
        self.db = db
        self._cfg: Optional[CommercialCodingConfig] = None

    # ==================================================================
    #  Configurazione codifica
    # ==================================================================

    def get_coding_config(self) -> CommercialCodingConfig:
        if self._cfg is None:
            self._cfg = self._load_config()
        return self._cfg

    def save_coding_config(self, cfg: CommercialCodingConfig):
        self._cfg = cfg
        with self.db.write_lock():
            with self.db.connection() as conn:
                conn.execute(
                    "INSERT INTO shared_settings(key, value) VALUES(?,?) "
                    "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                    ("commercial_coding_config", cfg.to_json()),
                )
                conn.commit()

    def _load_config(self) -> CommercialCodingConfig:
        row = self.db.fetchone(
            "SELECT value FROM shared_settings WHERE key='commercial_coding_config'"
        )
        if row and row.get("value"):
            try:
                return CommercialCodingConfig.from_json(row["value"])
            except Exception:
                pass
        return CommercialCodingConfig.default()

    # ==================================================================
    #  Categorie
    # ==================================================================

    def get_categories(self, item_type: str = None, only_active: bool = False) -> list:
        if item_type:
            return self.db.fetchall(
                "SELECT * FROM commercial_categories WHERE item_type=? ORDER BY code",
                (item_type,),
            )
        return self.db.fetchall("SELECT * FROM commercial_categories ORDER BY item_type, code")

    def get_category(self, cat_id: int) -> Optional[dict]:
        return self.db.fetchone(
            "SELECT * FROM commercial_categories WHERE id=?", (cat_id,)
        )

    def get_category_by_code(self, code: str, item_type: str = None) -> Optional[dict]:
        if item_type:
            return self.db.fetchone(
                "SELECT * FROM commercial_categories WHERE code=? AND item_type=?",
                (code.strip(), item_type),
            )
        return self.db.fetchone(
            "SELECT * FROM commercial_categories WHERE code=?", (code.strip(),)
        )

    def create_category(self, code: str, description: str,
                        item_type: str = "commerciale") -> int:
        return self.db.execute(
            "INSERT INTO commercial_categories(item_type, code, description) VALUES(?,?,?)",
            (item_type, code.strip(), description.strip()),
        )

    def update_category(self, cat_id: int, description: str):
        self.db.execute(
            "UPDATE commercial_categories SET description=? WHERE id=?",
            (description.strip(), cat_id),
        )

    def delete_category(self, cat_id: int):
        """Elimina una categoria (solo se priva di articoli associati)."""
        items = self.db.fetchone(
            "SELECT COUNT(*) AS n FROM commercial_items WHERE category_id=?", (cat_id,)
        )
        if items and items["n"] > 0:
            raise ValueError("Impossibile eliminare: esistono articoli in questa categoria.")
        self.db.execute(
            "DELETE FROM commercial_subcategories WHERE category_id=?", (cat_id,)
        )
        self.db.execute(
            "DELETE FROM commercial_categories WHERE id=?", (cat_id,)
        )

    def deactivate_category(self, cat_id: int):
        """Mantenuto per compatibilità — non più usato nell'UI."""
        pass

    def get_next_category_code(self, item_type: str) -> str:
        """Auto-genera il prossimo codice 4-cifre per una categoria del tipo dato."""
        rows = self.db.fetchall(
            "SELECT code FROM commercial_categories WHERE item_type=?", (item_type,)
        )
        existing = [r["code"] for r in rows]
        cfg = self.get_coding_config()
        return cfg.next_cat_code(existing)

    # ==================================================================
    #  Sottocategorie
    # ==================================================================

    def get_subcategories(self, category_id: int, only_active: bool = False) -> list:
        return self.db.fetchall(
            "SELECT * FROM commercial_subcategories WHERE category_id=? ORDER BY code",
            (category_id,),
        )

    def get_subcategory(self, sub_id: int) -> Optional[dict]:
        return self.db.fetchone(
            "SELECT * FROM commercial_subcategories WHERE id=?", (sub_id,)
        )

    def create_subcategory(self, category_id: int, code: str,
                           description: str, desc_template: str = "") -> int:
        return self.db.execute(
            "INSERT INTO commercial_subcategories"
            "(category_id, code, description, desc_template) VALUES(?,?,?,?)",
            (category_id, code.strip(), description.strip(), desc_template.strip()),
        )

    def update_subcategory(self, sub_id: int, description: str,
                           desc_template: str = ""):
        self.db.execute(
            "UPDATE commercial_subcategories "
            "SET description=?, desc_template=? WHERE id=?",
            (description.strip(), desc_template.strip(), sub_id),
        )

    def delete_subcategory(self, sub_id: int):
        """Elimina una sottocategoria (solo se priva di articoli associati)."""
        items = self.db.fetchone(
            "SELECT COUNT(*) AS n FROM commercial_items WHERE subcategory_id=?", (sub_id,)
        )
        if items and items["n"] > 0:
            raise ValueError("Impossibile eliminare: esistono articoli in questa sottocategoria.")
        self.db.execute("DELETE FROM commercial_subcategories WHERE id=?", (sub_id,))

    def deactivate_subcategory(self, sub_id: int):
        """Mantenuto per compatibilità — non più usato nell'UI."""
        pass

    def get_next_subcategory_code(self, category_id: int) -> str:
        """Auto-genera il prossimo codice 4-cifre per una sottocategoria della categoria data."""
        rows = self.db.fetchall(
            "SELECT code FROM commercial_subcategories WHERE category_id=?", (category_id,)
        )
        existing = [r["code"] for r in rows]
        cfg = self.get_coding_config()
        return cfg.next_sub_code(existing)

    # ==================================================================
    #  Codifica
    # ==================================================================

    def _type_prefix(self, item_type: str) -> str:
        """Ritorna il prefisso numerico per il tipo (es. '5' o '6')."""
        return COMMERCIAL_ITEM_TYPES.get(item_type, "5")

    def _peek_counter(self, item_type: str,
                      category_id: int,
                      subcategory_id: Optional[int]) -> int:
        """Legge il contatore senza incrementarlo."""
        row = self.db.fetchone(
            "SELECT last_value FROM commercial_counters "
            "WHERE item_type=? AND category_id=? AND subcategory_id IS ?",
            (item_type, category_id, subcategory_id),
        )
        return (row["last_value"] + 1) if row else 1

    def _increment_counter(self, item_type: str,
                           category_id: int,
                           subcategory_id: Optional[int]) -> int:
        """Incrementa il contatore e ritorna il nuovo valore."""
        with self.db.write_lock():
            with self.db.connection() as conn:
                row = conn.execute(
                    "SELECT id, last_value FROM commercial_counters "
                    "WHERE item_type=? AND category_id=? AND subcategory_id IS ?",
                    (item_type, category_id, subcategory_id),
                ).fetchone()
                if row:
                    new_val = row["last_value"] + 1
                    conn.execute(
                        "UPDATE commercial_counters SET last_value=? WHERE id=?",
                        (new_val, row["id"]),
                    )
                else:
                    new_val = 1
                    conn.execute(
                        "INSERT INTO commercial_counters"
                        "(item_type, category_id, subcategory_id, last_value) "
                        "VALUES(?,?,?,?)",
                        (item_type, category_id, subcategory_id, new_val),
                    )
                conn.commit()
                return new_val

    def _build_code(self, item_type: str, cat_code: str,
                    sub_code: Optional[str], num: int) -> str:
        cfg = self.get_coding_config()
        prefix_l = self._type_prefix(item_type)
        return cfg.render(prefix_l, cat_code, sub_code, num)

    def preview_code(self, item_type: str,
                     category_id: int,
                     subcategory_id: Optional[int] = None) -> str:
        """Restituisce il prossimo codice senza incrementare il contatore."""
        cat = self.get_category(category_id)
        if not cat:
            return "—"
        sub = self.get_subcategory(subcategory_id) if subcategory_id else None
        num = self._peek_counter(item_type, category_id, subcategory_id)
        return self._build_code(
            item_type, cat["code"],
            sub["code"] if sub else None, num
        )

    def next_code(self, item_type: str,
                  category_id: int,
                  subcategory_id: Optional[int] = None) -> str:
        """Incrementa il contatore e ritorna il nuovo codice."""
        cat = self.get_category(category_id)
        if not cat:
            raise ValueError("Categoria non trovata")
        sub = self.get_subcategory(subcategory_id) if subcategory_id else None
        num = self._increment_counter(item_type, category_id, subcategory_id)
        return self._build_code(
            item_type, cat["code"],
            sub["code"] if sub else None, num
        )

    # ==================================================================
    #  Articoli commerciali
    # ==================================================================

    def create_item(self, item_type: str, category_id: int,
                    subcategory_id: Optional[int],
                    description: str, notes: str = "",
                    created_by: Optional[int] = None) -> int:
        """Crea un nuovo articolo commerciale, generando il codice automaticamente."""
        code = self.next_code(item_type, category_id, subcategory_id)
        return self.db.execute(
            "INSERT INTO commercial_items"
            "(code, item_type, category_id, subcategory_id, description, "
            " notes, state, created_by) VALUES(?,?,?,?,?,?,'Attivo',?)",
            (code, item_type, category_id, subcategory_id,
             description.strip(), notes.strip(), created_by),
        )

    def update_item(self, item_id: int, description: str,
                    notes: str = "", modified_by: Optional[int] = None):
        self.db.execute(
            "UPDATE commercial_items "
            "SET description=?, notes=?, modified_by=?, modified_at=datetime('now') "
            "WHERE id=?",
            (description.strip(), notes.strip(), modified_by, item_id),
        )

    def duplicate_item(self, source_item_id: int,
                       created_by: Optional[int] = None) -> int:
        """
        Crea una copia dell'articolo con un nuovo codice generato automaticamente.
        Copia: item_type, category_id, subcategory_id, description, notes.
        Non copia: file SW, fornitori, proprietà SW.
        Ritorna il nuovo item_id.
        """
        src = self.get_item(source_item_id)
        if not src:
            raise ValueError(f"Articolo {source_item_id} non trovato.")
        return self.create_item(
            item_type=src["item_type"],
            category_id=src["category_id"],
            subcategory_id=src.get("subcategory_id"),
            description=src["description"],
            notes=src.get("notes") or "",
            created_by=created_by,
        )

    def get_item(self, item_id: int) -> Optional[dict]:
        return self.db.fetchone(
            """
            SELECT ci.*,
                   cc.code AS cat_code, cc.description AS cat_description,
                   cs.code AS sub_code, cs.description AS sub_description,
                   cs.desc_template,
                   u1.full_name AS created_by_name,
                   u2.full_name AS locked_by_name
            FROM commercial_items ci
            LEFT JOIN commercial_categories cc ON cc.id = ci.category_id
            LEFT JOIN commercial_subcategories cs ON cs.id = ci.subcategory_id
            LEFT JOIN users u1 ON u1.id = ci.created_by
            LEFT JOIN users u2 ON u2.id = ci.locked_by
            WHERE ci.id = ?
            """,
            (item_id,),
        )

    def get_item_by_code(self, code: str) -> Optional[dict]:
        return self.db.fetchone(
            "SELECT * FROM commercial_items WHERE code=?", (code,)
        )

    def search_items(self, text: str = "", category_id: int = 0,
                     subcategory_id: int = 0, supplier_id: int = 0,
                     state: str = "", item_type: str = "") -> list:
        """
        Ricerca articoli con filtri multipli.

        Returns lista di dict arricchiti con cat_code, sub_code,
        preferred_supplier_name.
        """
        conditions = []
        params: list = []

        if text:
            t = f"%{text.lower()}%"
            conditions.append(
                "(LOWER(ci.code) LIKE ? OR LOWER(ci.description) LIKE ? "
                " OR LOWER(ci.notes) LIKE ?)"
            )
            params.extend([t, t, t])

        if category_id:
            conditions.append("ci.category_id=?")
            params.append(category_id)

        if subcategory_id:
            conditions.append("ci.subcategory_id=?")
            params.append(subcategory_id)

        if state:
            conditions.append("ci.state=?")
            params.append(state)

        if item_type:
            conditions.append("ci.item_type=?")
            params.append(item_type)

        where = ("WHERE " + " AND ".join(conditions)) if conditions else ""

        # Filtro per fornitore (subquery)
        if supplier_id:
            where_clause = f"{where} AND " if where else "WHERE "
            where = (
                where_clause
                + "ci.id IN (SELECT item_id FROM commercial_item_suppliers "
                  "WHERE supplier_id=?)"
            )
            params.append(supplier_id)

        sql = f"""
            SELECT ci.*,
                   cc.code AS cat_code, cc.description AS cat_description,
                   cs.code AS sub_code, cs.description AS sub_description,
                   u1.full_name AS created_by_name,
                   u2.full_name AS locked_by_name,
                   (SELECT s.name
                    FROM commercial_item_suppliers cis2
                    JOIN commercial_suppliers s ON s.id = cis2.supplier_id
                    WHERE cis2.item_id = ci.id AND cis2.is_preferred = 1
                    LIMIT 1) AS preferred_supplier_name
            FROM commercial_items ci
            LEFT JOIN commercial_categories cc ON cc.id = ci.category_id
            LEFT JOIN commercial_subcategories cs ON cs.id = ci.subcategory_id
            LEFT JOIN users u1 ON u1.id = ci.created_by
            LEFT JOIN users u2 ON u2.id = ci.locked_by
            {where}
            ORDER BY ci.item_type, cc.code, cs.code, ci.code
        """
        return self.db.fetchall(sql, params)

    def delete_item(self, item_id: int):
        """Elimina articolo se non è bloccato."""
        item = self.get_item(item_id)
        if not item:
            raise ValueError("Articolo non trovato")
        if item["is_locked"]:
            raise PermissionError("Articolo in checkout: impossibile eliminare")
        # Elimina cascading manuale (SQLite non ha CASCADE di default qui)
        self.db.execute(
            "DELETE FROM commercial_item_suppliers WHERE item_id=?", (item_id,)
        )
        self.db.execute(
            "DELETE FROM commercial_properties WHERE item_id=?", (item_id,)
        )
        self.db.execute(
            "DELETE FROM commercial_checkout_log WHERE item_id=?", (item_id,)
        )
        self.db.execute(
            "DELETE FROM asm_commercial_components WHERE child_commercial_id=?",
            (item_id,),
        )
        self.db.execute(
            "DELETE FROM commercial_items WHERE id=?", (item_id,)
        )

    # ==================================================================
    #  Workflow stati
    # ==================================================================

    def change_state(self, item_id: int, to_state: str,
                     user_id: int, notes: str = ""):
        item = self.get_item(item_id)
        if not item:
            raise ValueError("Articolo non trovato")
        current = item["state"]
        allowed = COMMERCIAL_WORKFLOW_TRANSITIONS.get(current, [])
        if to_state not in allowed:
            raise ValueError(
                f"Transizione '{current}' → '{to_state}' non consentita"
            )
        self.db.execute(
            "UPDATE commercial_items SET state=?, modified_by=?, "
            "modified_at=datetime('now') WHERE id=?",
            (to_state, user_id, item_id),
        )

    # ==================================================================
    #  Fornitori
    # ==================================================================

    def get_suppliers(self, only_active: bool = True) -> list:
        sql = "SELECT * FROM commercial_suppliers"
        if only_active:
            sql += " WHERE active=1"
        sql += " ORDER BY name"
        return self.db.fetchall(sql)

    def get_supplier(self, sup_id: int) -> Optional[dict]:
        return self.db.fetchone(
            "SELECT * FROM commercial_suppliers WHERE id=?", (sup_id,)
        )

    def create_supplier(self, name: str, short_code: str = "",
                        contact: str = "", email: str = "",
                        phone: str = "", website: str = "",
                        notes: str = "") -> int:
        return self.db.execute(
            "INSERT INTO commercial_suppliers"
            "(name, short_code, contact, email, phone, website, notes) "
            "VALUES(?,?,?,?,?,?,?)",
            (name.strip(), short_code.strip(), contact.strip(),
             email.strip(), phone.strip(), website.strip(), notes.strip()),
        )

    def update_supplier(self, sup_id: int, name: str, short_code: str = "",
                        contact: str = "", email: str = "", phone: str = "",
                        website: str = "", notes: str = "",
                        active: int = 1):
        self.db.execute(
            "UPDATE commercial_suppliers "
            "SET name=?, short_code=?, contact=?, email=?, phone=?, "
            "website=?, notes=?, active=? WHERE id=?",
            (name.strip(), short_code.strip(), contact.strip(),
             email.strip(), phone.strip(), website.strip(),
             notes.strip(), active, sup_id),
        )

    def deactivate_supplier(self, sup_id: int):
        self.db.execute(
            "UPDATE commercial_suppliers SET active=0 WHERE id=?", (sup_id,)
        )

    def get_item_suppliers(self, item_id: int) -> list:
        return self.db.fetchall(
            """
            SELECT cis.*, s.name AS supplier_name, s.short_code
            FROM commercial_item_suppliers cis
            JOIN commercial_suppliers s ON s.id = cis.supplier_id
            WHERE cis.item_id = ?
            ORDER BY cis.is_preferred DESC, s.name
            """,
            (item_id,),
        )

    def add_item_supplier(self, item_id: int, supplier_id: int,
                          supplier_code: str = "", unit_price: Optional[float] = None,
                          currency: str = "EUR", lead_time_days: Optional[int] = None,
                          is_preferred: bool = False, notes: str = "") -> int:
        return self.db.execute(
            "INSERT INTO commercial_item_suppliers"
            "(item_id, supplier_id, supplier_code, unit_price, currency, "
            " lead_time_days, is_preferred, notes) VALUES(?,?,?,?,?,?,?,?)",
            (item_id, supplier_id, supplier_code.strip(),
             unit_price, currency, lead_time_days,
             1 if is_preferred else 0, notes.strip()),
        )

    def update_item_supplier(self, link_id: int, supplier_code: str = "",
                             unit_price: Optional[float] = None,
                             currency: str = "EUR",
                             lead_time_days: Optional[int] = None,
                             is_preferred: bool = False,
                             notes: str = ""):
        self.db.execute(
            "UPDATE commercial_item_suppliers "
            "SET supplier_code=?, unit_price=?, currency=?, lead_time_days=?, "
            "is_preferred=?, notes=?, updated_at=datetime('now') WHERE id=?",
            (supplier_code.strip(), unit_price, currency,
             lead_time_days, 1 if is_preferred else 0,
             notes.strip(), link_id),
        )

    def remove_item_supplier(self, link_id: int):
        self.db.execute(
            "DELETE FROM commercial_item_suppliers WHERE id=?", (link_id,)
        )

    def set_preferred_supplier(self, item_id: int, supplier_id: int):
        """Imposta un fornitore come preferito, deselezionando gli altri."""
        with self.db.write_lock():
            with self.db.connection() as conn:
                conn.execute(
                    "UPDATE commercial_item_suppliers SET is_preferred=0 "
                    "WHERE item_id=?",
                    (item_id,),
                )
                conn.execute(
                    "UPDATE commercial_item_suppliers SET is_preferred=1 "
                    "WHERE item_id=? AND supplier_id=?",
                    (item_id, supplier_id),
                )
                conn.commit()

    # ==================================================================
    #  Proprietà SolidWorks
    # ==================================================================

    def save_properties(self, item_id: int, props: dict):
        """Salva o aggiorna le proprietà personalizzate di un articolo."""
        with self.db.write_lock():
            with self.db.connection() as conn:
                for name, value in props.items():
                    conn.execute(
                        "INSERT INTO commercial_properties(item_id, prop_name, prop_value) "
                        "VALUES(?,?,?) ON CONFLICT(item_id, prop_name) "
                        "DO UPDATE SET prop_value=excluded.prop_value",
                        (item_id, name, str(value) if value is not None else None),
                    )
                conn.commit()

    def get_properties(self, item_id: int) -> dict:
        rows = self.db.fetchall(
            "SELECT prop_name, prop_value FROM commercial_properties WHERE item_id=?",
            (item_id,),
        )
        return {r["prop_name"]: r["prop_value"] for r in rows}

    def sync_sw_to_pdm(self, item_id: int, file_path: Path,
                       properties_manager) -> dict:
        """
        Legge le proprietà dal file SolidWorks e le salva in commercial_properties.

        Args:
            properties_manager: istanza di PropertiesManager (per la lettura COM)
        """
        try:
            props = properties_manager.read_from_sw_file(file_path)
        except Exception as e:
            return {"ok": False, "error": str(e), "imported_count": 0}
        self.save_properties(item_id, props)
        return {"ok": True, "imported_count": len(props)}

    def sync_pdm_to_sw(self, item_id: int, file_path: Path,
                       properties_manager) -> dict:
        """
        Scrive le proprietà PDM nel file SolidWorks via COM API.

        Args:
            properties_manager: istanza di PropertiesManager
        """
        item = self.get_item(item_id)
        if not item:
            return {"ok": False, "error": "Articolo non trovato", "written_count": 0}
        # Proprietà base del PDM da esportare
        pdm_props = {
            "Code":        item["code"],
            "Description": item["description"],
            "State":       item["state"],
            "ItemType":    item["item_type"],
        }
        # Aggiungi proprietà custom
        pdm_props.update(self.get_properties(item_id))
        try:
            properties_manager.write_to_sw_file(file_path, pdm_props)
        except Exception as e:
            return {"ok": False, "error": str(e), "written_count": 0}
        return {"ok": True, "written_count": len(pdm_props)}

    # ==================================================================
    #  BOM CAD ↔ Articoli commerciali
    # ==================================================================

    def get_commercial_bom(self, parent_doc_id: int) -> list:
        """Articoli commerciali presenti nella BOM di un assieme CAD."""
        return self.db.fetchall(
            """
            SELECT acc.*, ci.code, ci.description, ci.item_type, ci.state,
                   cc.description AS cat_description,
                   cs.description AS sub_description,
                   (SELECT s.name
                    FROM commercial_item_suppliers cis
                    JOIN commercial_suppliers s ON s.id = cis.supplier_id
                    WHERE cis.item_id = ci.id AND cis.is_preferred = 1
                    LIMIT 1) AS preferred_supplier
            FROM asm_commercial_components acc
            JOIN commercial_items ci ON ci.id = acc.child_commercial_id
            LEFT JOIN commercial_categories cc ON cc.id = ci.category_id
            LEFT JOIN commercial_subcategories cs ON cs.id = ci.subcategory_id
            WHERE acc.parent_doc_id = ?
            ORDER BY acc.position, ci.code
            """,
            (parent_doc_id,),
        )

    def add_to_bom(self, parent_doc_id: int, item_id: int,
                   quantity: float = 1.0,
                   position: str = "", notes: str = "") -> int:
        return self.db.execute(
            "INSERT INTO asm_commercial_components"
            "(parent_doc_id, child_commercial_id, quantity, position, notes) "
            "VALUES(?,?,?,?,?)",
            (parent_doc_id, item_id, quantity, position.strip(), notes.strip()),
        )

    def remove_from_bom(self, link_id: int):
        self.db.execute(
            "DELETE FROM asm_commercial_components WHERE id=?", (link_id,)
        )

    def update_bom_row(self, link_id: int, quantity: float,
                       position: str = "", notes: str = ""):
        self.db.execute(
            "UPDATE asm_commercial_components "
            "SET quantity=?, position=?, notes=? WHERE id=?",
            (quantity, position.strip(), notes.strip(), link_id),
        )

    # ==================================================================
    #  Checkout / Checkin (file SolidWorks opzionale)
    # ==================================================================

    @staticmethod
    def _set_readonly(path: Path):
        try:
            path.chmod(
                path.stat().st_mode
                & ~(stat_mod.S_IWRITE | stat_mod.S_IWGRP | stat_mod.S_IWOTH)
            )
        except OSError:
            pass

    @staticmethod
    def _set_writable(path: Path):
        try:
            path.chmod(path.stat().st_mode | stat_mod.S_IWRITE)
        except OSError:
            pass

    @staticmethod
    def _md5(filepath: Path) -> str:
        h = hashlib.md5()
        with open(filepath, "rb") as f:
            for chunk in iter(lambda: f.read(8192), b""):
                h.update(chunk)
        return h.hexdigest()

    @staticmethod
    def _file_snapshot(filepath: Path) -> dict:
        if not filepath or not filepath.exists():
            return {"md5": "", "size": 0, "mtime": 0.0}
        s = filepath.stat()
        return {
            "md5":   CommercialManager._md5(filepath),
            "size":  s.st_size,
            "mtime": s.st_mtime,
        }

    def _commercial_archive_root(self, shared_paths: "SharedPaths") -> Path:
        """
        Ritorna la cartella radice dell'archivio file SW commerciali.
        Se configurato in CommercialCodingConfig.commercial_archive_path, usa quello.
        Altrimenti usa il default {SharedPaths.archive}/commercial/.
        """
        cfg = self.get_coding_config()
        p = (cfg.commercial_archive_path or "").strip()
        if p:
            pp = Path(p)
            return pp if pp.is_absolute() else shared_paths.root / pp
        return shared_paths.archive / "commercial"

    def _archive_file_path(self, item: dict,
                           shared_paths: "SharedPaths") -> Optional[Path]:
        if item.get("archive_path"):
            return shared_paths.root / item["archive_path"]
        if item.get("file_name") and item.get("file_ext"):
            p = self._commercial_archive_root(shared_paths) / item["code"]
            return p / (item["code"] + item["file_ext"])
        return None

    def _ws_file_path(self, item: dict,
                      shared_paths: "SharedPaths") -> Path:
        from config import load_local_config
        cfg = load_local_config()
        ws = Path(cfg.get("sw_workspace", ""))
        ext = item.get("file_ext") or ".SLDPRT"
        return ws / (item["code"] + ext)

    def checkout_item(self, item_id: int, current_user: dict,
                      shared_paths: "SharedPaths") -> Path:
        """
        Esegue il checkout di un articolo commerciale con file SW collegato.

        Returns:
            Path al file nella workspace locale.
        Raises:
            PermissionError / RuntimeError se non consentito.
        """
        item = self.get_item(item_id)
        if not item:
            raise ValueError("Articolo non trovato")
        if item["state"] in COMMERCIAL_READONLY_STATES:
            raise PermissionError(
                f"L'articolo è '{item['state']}': checkout non consentito"
            )
        if item["is_locked"]:
            raise PermissionError(
                f"Articolo già in checkout da: {item.get('locked_by_name', '?')}"
            )
        if not item.get("file_name"):
            raise ValueError("Nessun file SolidWorks collegato all'articolo")

        archive_path = self._archive_file_path(item, shared_paths)
        if not archive_path or not archive_path.exists():
            raise FileNotFoundError(
                f"File archiviato non trovato: {archive_path}"
            )

        ws_path = self._ws_file_path(item, shared_paths)
        ws_path.parent.mkdir(parents=True, exist_ok=True)

        # Copia archive → workspace
        shutil.copy2(archive_path, ws_path)
        self._set_writable(ws_path)

        snap = self._file_snapshot(archive_path)
        ws_name = socket.gethostname()
        user_id = current_user["id"]

        with self.db.write_lock():
            with self.db.connection() as conn:
                conn.execute(
                    "UPDATE commercial_items SET is_locked=1, locked_by=?, "
                    "locked_at=datetime('now'), locked_ws=?, "
                    "checkout_md5=?, checkout_size=?, checkout_mtime=? "
                    "WHERE id=?",
                    (user_id, ws_name,
                     snap["md5"], snap["size"], snap["mtime"],
                     item_id),
                )
                conn.execute(
                    "INSERT INTO commercial_checkout_log"
                    "(item_id, user_id, action, workstation, workspace_path, "
                    " checkout_md5, checkout_size, checkout_mtime) "
                    "VALUES(?,?,?,?,?,?,?,?)",
                    (item_id, user_id, "checkout", ws_name, str(ws_path),
                     snap["md5"], snap["size"], snap["mtime"]),
                )
                conn.commit()

        return ws_path

    def checkin_item(self, item_id: int, current_user: dict,
                     shared_paths: "SharedPaths",
                     archive_file: bool = True,
                     notes: str = "") -> dict:
        """
        Esegue il checkin di un articolo commerciale.

        Returns:
            dict con: modified, conflict, archived
        """
        item = self.get_item(item_id)
        if not item:
            raise ValueError("Articolo non trovato")

        user_id = current_user["id"]
        if item["locked_by"] != user_id:
            raise PermissionError(
                "Il checkout appartiene a un altro utente"
            )
        if item["state"] in COMMERCIAL_READONLY_STATES:
            raise PermissionError(
                f"L'articolo è '{item['state']}': checkin non consentito"
            )

        ws_path = self._ws_file_path(item, shared_paths)
        archive_path = self._archive_file_path(item, shared_paths)

        modified  = False
        conflict  = False
        archived  = False

        if ws_path.exists() and archive_path:
            ws_md5      = self._md5(ws_path)
            checkout_md5 = item.get("checkout_md5") or ""
            cur_archive_md5 = (
                self._md5(archive_path) if archive_path.exists() else ""
            )

            modified = (ws_md5 != checkout_md5)
            conflict = (
                cur_archive_md5 != checkout_md5 and cur_archive_md5 != ""
            )

            if archive_file:
                archive_path.parent.mkdir(parents=True, exist_ok=True)
                if archive_path.exists():
                    self._set_writable(archive_path)
                shutil.copy2(ws_path, archive_path)
                self._set_readonly(archive_path)
                archived = True

        ws_name = socket.gethostname()
        with self.db.write_lock():
            with self.db.connection() as conn:
                conn.execute(
                    "UPDATE commercial_items SET is_locked=0, locked_by=NULL, "
                    "locked_at=NULL, locked_ws=NULL, "
                    "checkout_md5=NULL, checkout_size=NULL, checkout_mtime=NULL, "
                    "modified_by=?, modified_at=datetime('now') WHERE id=?",
                    (user_id, item_id),
                )
                conn.execute(
                    "INSERT INTO commercial_checkout_log"
                    "(item_id, user_id, action, workstation, workspace_path, notes) "
                    "VALUES(?,?,?,?,?,?)",
                    (item_id, user_id, "checkin", ws_name, str(ws_path), notes),
                )
                conn.commit()

        return {"modified": modified, "conflict": conflict, "archived": archived}

    def undo_checkout_item(self, item_id: int, current_user: dict) -> bool:
        """Annulla il checkout senza archiviare il file."""
        item = self.get_item(item_id)
        if not item:
            return False
        user_id = current_user["id"]
        is_admin = current_user.get("role") == "Amministratore"
        if item["locked_by"] != user_id and not is_admin:
            raise PermissionError(
                "Solo il proprietario del checkout o un amministratore "
                "può annullare il checkout"
            )
        ws_name = socket.gethostname()
        with self.db.write_lock():
            with self.db.connection() as conn:
                conn.execute(
                    "UPDATE commercial_items SET is_locked=0, locked_by=NULL, "
                    "locked_at=NULL, locked_ws=NULL, "
                    "checkout_md5=NULL, checkout_size=NULL, checkout_mtime=NULL "
                    "WHERE id=?",
                    (item_id,),
                )
                conn.execute(
                    "INSERT INTO commercial_checkout_log"
                    "(item_id, user_id, action, workstation) VALUES(?,?,?,?)",
                    (item_id, user_id, "undo_checkout", ws_name),
                )
                conn.commit()
        return True

    # ==================================================================
    #  Collegamento file SolidWorks
    # ==================================================================

    def link_sw_file(self, item_id: int, file_path: Path,
                     shared_paths: "SharedPaths",
                     current_user: Optional[dict] = None):
        """
        Collega un file SolidWorks all'articolo e lo copia in archivio.

        Il file viene copiato da workspace → archive/commercial/{code}/
        """
        item = self.get_item(item_id)
        if not item:
            raise ValueError("Articolo non trovato")

        code = item["code"]
        ext  = file_path.suffix.upper()
        dest_dir = self._commercial_archive_root(shared_paths) / code
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest = dest_dir / (code + ext)

        if dest.exists():
            self._set_writable(dest)
        shutil.copy2(file_path, dest)
        self._set_readonly(dest)

        rel_path = str(dest.relative_to(shared_paths.root))
        user_id = current_user["id"] if current_user else None

        self.db.execute(
            "UPDATE commercial_items "
            "SET file_name=?, file_ext=?, archive_path=?, "
            "modified_by=?, modified_at=datetime('now') WHERE id=?",
            (code + ext, ext, rel_path, user_id, item_id),
        )

    def unlink_sw_file(self, item_id: int, current_user: Optional[dict] = None):
        """Rimuove il collegamento al file SolidWorks (non elimina il file)."""
        user_id = current_user["id"] if current_user else None
        self.db.execute(
            "UPDATE commercial_items "
            "SET file_name=NULL, file_ext=NULL, archive_path=NULL, "
            "thumbnail=NULL, modified_by=?, modified_at=datetime('now') "
            "WHERE id=?",
            (user_id, item_id),
        )

    # ==================================================================
    #  Log checkout (lettura)
    # ==================================================================

    def get_checkout_log(self, item_id: int) -> list:
        return self.db.fetchall(
            """
            SELECT cl.*, u.full_name AS user_name
            FROM commercial_checkout_log cl
            LEFT JOIN users u ON u.id = cl.user_id
            WHERE cl.item_id = ?
            ORDER BY cl.timestamp DESC
            """,
            (item_id,),
        )

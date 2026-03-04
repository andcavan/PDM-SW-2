# =============================================================================
#  core/coding_manager.py  –  Gestione codifica gerarchica documenti
#
#  Schema codici:
#    LIV0 – Macchina (ASM) :  MMM_V001
#    LIV1 – Gruppo   (ASM) :  MMM_GGGG-V001
#    LIV2 – Parte    (PRT) :  MMM_GGGG-0001  (sale da 0001)
#    LIV2 – Sottogruppo (ASM): MMM_GGGG-9999 (scende da 9999)
#
#  Separatori: '_' tra MMM e GGGG, '-' tra GGGG e numero
#  Versione:   indipendente per ogni macchina (LIV0) e per ogni
#              coppia macchina+gruppo (LIV1)
#  Collisione: warning quando parte e sottogruppo si avvicinano
# =============================================================================
from __future__ import annotations
import re
import warnings
from typing import Optional, TYPE_CHECKING

if TYPE_CHECKING:
    from core.database import Database

# Soglia warning collisione LIV2
COLLISION_WARNING_THRESHOLD = 500
# Numeri LIV2: parti salgono da 1 a PART_MAX, sottogruppi scendono da 9999
PART_MAX    = 8999
SUBGR_MIN   = 9000
SUBGR_START = 9999


class CodingManager:
    def __init__(self, db: "Database"):
        self.db = db

    # ==================================================================
    # MACCHINE
    # ==================================================================

    def create_machine(self, code: str, description: str,
                       code_type: str = "ALPHA", code_length: int = 3) -> int:
        """Crea una nuova macchina. Ritorna l'id."""
        code = code.upper().strip()
        self._validate_code(code, code_type, code_length)
        return self.db.execute(
            """INSERT INTO machines (code, description, code_type, code_length)
               VALUES (?,?,?,?)""",
            (code, description, code_type, code_length),
        )

    def update_machine(self, machine_id: int, description: str,
                       code_type: str, code_length: int):
        self.db.execute(
            """UPDATE machines SET description=?, code_type=?, code_length=?
               WHERE id=?""",
            (description, code_type, code_length, machine_id),
        )

    def deactivate_machine(self, machine_id: int):
        self.db.execute(
            "UPDATE machines SET active=0 WHERE id=?", (machine_id,)
        )

    def get_machines(self, only_active: bool = True) -> list:
        where = "WHERE active=1" if only_active else ""
        return self.db.fetchall(
            f"SELECT * FROM machines {where} ORDER BY code"
        )

    def get_machine(self, machine_id: int) -> Optional[dict]:
        return self.db.fetchone(
            "SELECT * FROM machines WHERE id=?", (machine_id,)
        )

    def get_machine_by_code(self, code: str) -> Optional[dict]:
        return self.db.fetchone(
            "SELECT * FROM machines WHERE code=?", (code.upper(),)
        )

    # ==================================================================
    # GRUPPI
    # ==================================================================

    def create_group(self, machine_id: int, code: str, description: str,
                     code_type: str = "ALPHA", code_length: int = 4) -> int:
        """Crea un nuovo gruppo per la macchina specificata."""
        code = code.upper().strip()
        self._validate_code(code, code_type, code_length)
        return self.db.execute(
            """INSERT INTO machine_groups
               (machine_id, code, description, code_type, code_length)
               VALUES (?,?,?,?,?)""",
            (machine_id, code, description, code_type, code_length),
        )

    def update_group(self, group_id: int, description: str,
                     code_type: str, code_length: int):
        self.db.execute(
            """UPDATE machine_groups SET description=?, code_type=?, code_length=?
               WHERE id=?""",
            (description, code_type, code_length, group_id),
        )

    def deactivate_group(self, group_id: int):
        self.db.execute(
            "UPDATE machine_groups SET active=0 WHERE id=?", (group_id,)
        )

    def get_groups(self, machine_id: int, only_active: bool = True) -> list:
        where = "AND active=1" if only_active else ""
        return self.db.fetchall(
            f"""SELECT * FROM machine_groups
                WHERE machine_id=? {where} ORDER BY code""",
            (machine_id,),
        )

    def get_group(self, group_id: int) -> Optional[dict]:
        return self.db.fetchone(
            "SELECT * FROM machine_groups WHERE id=?", (group_id,)
        )

    def get_all_groups(self, only_active: bool = True) -> list:
        where = "WHERE mg.active=1" if only_active else ""
        return self.db.fetchall(
            f"""SELECT mg.*, m.code as machine_code, m.description as machine_desc
                FROM machine_groups mg
                JOIN machines m ON m.id = mg.machine_id
                {where}
                ORDER BY m.code, mg.code"""
        )

    # ==================================================================
    # GENERAZIONE CODICI
    # ==================================================================

    def next_code_liv0(self, machine_id: int) -> str:
        """LIV0 – Macchina ASM: MMM_V001"""
        machine = self._get_machine_or_raise(machine_id)
        val = self._increment_counter("VERSION", machine_id, None)
        return f"{machine['code']}_V{str(val).zfill(3)}"

    def next_code_liv1(self, machine_id: int, group_id: int) -> str:
        """LIV1 – Gruppo ASM: MMM_GGGG-V001"""
        machine = self._get_machine_or_raise(machine_id)
        group   = self._get_group_or_raise(group_id)
        val = self._increment_counter("VERSION", machine_id, group_id)
        return f"{machine['code']}_{group['code']}-V{str(val).zfill(3)}"

    def next_code_liv2_part(self, machine_id: int, group_id: int) -> str:
        """LIV2 – Parte PRT: MMM_GGGG-0001 (sale)"""
        machine = self._get_machine_or_raise(machine_id)
        group   = self._get_group_or_raise(group_id)
        val = self._increment_counter("PART", machine_id, group_id)
        if val > PART_MAX:
            raise ValueError(
                f"Contatore parti esaurito per {machine['code']}_{group['code']} "
                f"(massimo {PART_MAX})"
            )
        self._check_collision(machine_id, group_id)
        return f"{machine['code']}_{group['code']}-{str(val).zfill(4)}"

    def next_code_liv2_subgroup(self, machine_id: int, group_id: int) -> str:
        """LIV2 – Sottogruppo ASM: MMM_GGGG-9999 (scende)"""
        machine = self._get_machine_or_raise(machine_id)
        group   = self._get_group_or_raise(group_id)
        val = self._decrement_counter("SUBGROUP", machine_id, group_id)
        if val < SUBGR_MIN:
            raise ValueError(
                f"Contatore sottogruppi esaurito per "
                f"{machine['code']}_{group['code']} (minimo {SUBGR_MIN})"
            )
        self._check_collision(machine_id, group_id)
        return f"{machine['code']}_{group['code']}-{str(val).zfill(4)}"

    # ------------------------------------------------------------------
    # Anteprima senza modificare il contatore
    # ------------------------------------------------------------------
    def preview_code(self, level: int, doc_subtype: str,
                     machine_id: int, group_id: int = None) -> str:
        """
        Ritorna il prossimo codice SENZA modificare il contatore.
        doc_subtype: 'ASM' per LIV0/LIV1/Sottogruppo, 'PRT' per Parte LIV2.
        """
        machine = self.get_machine(machine_id)
        if not machine:
            return "—"

        if level == 0:
            val = self._peek_counter("VERSION", machine_id, None) + 1
            return f"{machine['code']}_V{str(val).zfill(3)}"

        group = self.get_group(group_id) if group_id else None
        if not group:
            return f"{machine['code']}_???-???"

        if level == 1:
            val = self._peek_counter("VERSION", machine_id, group_id) + 1
            return f"{machine['code']}_{group['code']}-V{str(val).zfill(3)}"

        # LIV2
        if doc_subtype == "PRT":
            val = self._peek_counter("PART", machine_id, group_id) + 1
            return f"{machine['code']}_{group['code']}-{str(val).zfill(4)}"
        else:
            cur = self._peek_counter("SUBGROUP", machine_id, group_id)
            val = SUBGR_START if cur == 0 else cur - 1
            return f"{machine['code']}_{group['code']}-{str(val).zfill(4)}"

    # ==================================================================
    # CONTATORI
    # ==================================================================

    def get_counters(self, machine_id: int = None, group_id: int = None) -> list:
        """Ritorna i contatori, opzionalmente filtrati."""
        conditions, params = ["1=1"], []
        if machine_id is not None:
            conditions.append("hc.machine_id=?")
            params.append(machine_id)
        if group_id is not None:
            conditions.append("hc.group_id=?")
            params.append(group_id)
        return self.db.fetchall(
            f"""SELECT hc.*,
                       m.code  as machine_code,
                       mg.code as group_code
                FROM hierarchical_counters hc
                LEFT JOIN machines m ON m.id = hc.machine_id
                LEFT JOIN machine_groups mg ON mg.id = hc.group_id
                WHERE {' AND '.join(conditions)}
                ORDER BY hc.counter_type, m.code, mg.code""",
            params,
        )

    def reset_counter(self, counter_type: str,
                      machine_id: int, group_id: Optional[int],
                      value: int = 0):
        """Azzera o imposta manualmente un contatore."""
        row = self.db.fetchone(
            """SELECT id FROM hierarchical_counters
               WHERE counter_type=? AND machine_id IS ? AND group_id IS ?""",
            (counter_type, machine_id, group_id),
        )
        if row:
            self.db.execute(
                "UPDATE hierarchical_counters SET last_value=? WHERE id=?",
                (value, row["id"]),
            )
        else:
            self.db.execute(
                """INSERT INTO hierarchical_counters
                   (counter_type, machine_id, group_id, last_value)
                   VALUES (?,?,?,?)""",
                (counter_type, machine_id, group_id, value),
            )

    def get_collision_status(self, machine_id: int, group_id: int) -> dict:
        """
        Ritorna lo stato dei contatori LIV2 per un gruppo.
        Usato per mostrare indicatore visuale di spazio disponibile.
        """
        part_val  = self._peek_counter("PART",     machine_id, group_id)
        subgr_val = self._peek_counter("SUBGROUP", machine_id, group_id)
        if subgr_val == 0:
            subgr_val = SUBGR_START + 1   # non ancora usato

        gap       = subgr_val - part_val - 1
        warning   = gap < COLLISION_WARNING_THRESHOLD
        exhausted = gap <= 0
        return {
            "part_last":     part_val,
            "subgroup_last": subgr_val if (subgr_val != SUBGR_START + 1) else None,
            "gap":           gap,
            "warning":       warning,
            "exhausted":     exhausted,
        }

    # ==================================================================
    # VALIDAZIONE / PARSING
    # ==================================================================

    def validate_code_string(self, code: str, code_type: str,
                              code_length: int) -> tuple:
        """Verifica formato del codice macchina/gruppo."""
        code = code.upper().strip()
        if len(code) != code_length:
            return False, f"Lunghezza deve essere {code_length} caratteri"
        if code_type == "ALPHA":
            if not re.match(r'^[A-Z]+$', code):
                return False, "Solo lettere maiuscole (A-Z)"
        elif code_type == "NUM":
            if not re.match(r'^[0-9]+$', code):
                return False, "Solo cifre (0-9)"
        else:
            return False, f"Tipo non valido: {code_type}"
        return True, ""

    def parse_code(self, code: str) -> Optional[dict]:
        """
        Parsa un codice esistente e ritorna il dizionario con le parti.
        Ritorna None se il formato non è riconosciuto.
        """
        # LIV0: MMM_V001
        m = re.match(r'^([A-Z0-9]+)_V(\d+)$', code)
        if m:
            return {"level": 0, "machine": m.group(1),
                    "group": None, "subtype": "ASM",
                    "number": int(m.group(2))}

        # LIV1: MMM_GGGG-V001
        m = re.match(r'^([A-Z0-9]+)_([A-Z0-9]+)-V(\d+)$', code)
        if m:
            return {"level": 1, "machine": m.group(1),
                    "group": m.group(2), "subtype": "ASM",
                    "number": int(m.group(3))}

        # LIV2: MMM_GGGG-0001 oppure MMM_GGGG-9999
        m = re.match(r'^([A-Z0-9]+)_([A-Z0-9]+)-(\d{4})$', code)
        if m:
            num = int(m.group(3))
            subtype = "ASM" if num >= SUBGR_MIN else "PRT"
            return {"level": 2, "machine": m.group(1),
                    "group": m.group(2), "subtype": subtype,
                    "number": num}

        return None

    # ==================================================================
    # Retrocompatibilità (usata da altri moduli / CodingDialog legacy)
    # ==================================================================

    def get_config(self, doc_type: str) -> dict:
        """Stub per retrocompatibilità – ritorna dizionario vuoto."""
        return {}

    def save_config(self, *args, **kwargs):
        """Stub per retrocompatibilità."""
        pass

    def is_code_available(self, code: str, revision: str = "00",
                           doc_type: str = None) -> bool:
        """Ritorna True se la coppia (code, revision) è libera.
        Se doc_type è fornito, controlla unicità per (code, revision, doc_type):
        questo permette a un Disegno di avere lo stesso codice del PRT/ASM padre.
        """
        if doc_type:
            row = self.db.fetchone(
                "SELECT id FROM documents WHERE code=? AND revision=? AND doc_type=?",
                (code, revision, doc_type),
            )
        else:
            row = self.db.fetchone(
                "SELECT id FROM documents WHERE code=? AND revision=?",
                (code, revision),
            )
        return row is None

    def validate_custom_code(self, code: str) -> bool:
        return bool(re.match(r'^[A-Za-z0-9\-_\.]+$', code))

    def get_all_configs(self) -> list:
        """Stub per retrocompatibilità."""
        return []

    # ==================================================================
    # Metodi privati
    # ==================================================================

    def _get_machine_or_raise(self, machine_id: int) -> dict:
        m = self.get_machine(machine_id)
        if not m:
            raise ValueError(f"Macchina id={machine_id} non trovata")
        if not m["active"]:
            raise ValueError(f"Macchina '{m['code']}' non attiva")
        return m

    def _get_group_or_raise(self, group_id: int) -> dict:
        g = self.get_group(group_id)
        if not g:
            raise ValueError(f"Gruppo id={group_id} non trovato")
        if not g["active"]:
            raise ValueError(f"Gruppo '{g['code']}' non attivo")
        return g

    def _peek_counter(self, counter_type: str,
                      machine_id: Optional[int],
                      group_id: Optional[int]) -> int:
        """Legge il valore corrente senza modificarlo."""
        row = self.db.fetchone(
            """SELECT last_value FROM hierarchical_counters
               WHERE counter_type=? AND machine_id IS ? AND group_id IS ?""",
            (counter_type, machine_id, group_id),
        )
        return row["last_value"] if row else 0

    def _increment_counter(self, counter_type: str,
                           machine_id: Optional[int],
                           group_id: Optional[int]) -> int:
        cur = self._peek_counter(counter_type, machine_id, group_id)
        new_val = cur + 1
        self._write_counter(counter_type, machine_id, group_id, new_val)
        return new_val

    def _decrement_counter(self, counter_type: str,
                           machine_id: Optional[int],
                           group_id: Optional[int]) -> int:
        cur = self._peek_counter(counter_type, machine_id, group_id)
        new_val = SUBGR_START if cur == 0 else cur - 1
        self._write_counter(counter_type, machine_id, group_id, new_val)
        return new_val

    def _write_counter(self, counter_type: str,
                       machine_id: Optional[int],
                       group_id: Optional[int], value: int):
        row = self.db.fetchone(
            """SELECT id FROM hierarchical_counters
               WHERE counter_type=? AND machine_id IS ? AND group_id IS ?""",
            (counter_type, machine_id, group_id),
        )
        if row:
            self.db.execute(
                "UPDATE hierarchical_counters SET last_value=? WHERE id=?",
                (value, row["id"]),
            )
        else:
            self.db.execute(
                """INSERT INTO hierarchical_counters
                   (counter_type, machine_id, group_id, last_value)
                   VALUES (?,?,?,?)""",
                (counter_type, machine_id, group_id, value),
            )

    def _check_collision(self, machine_id: int, group_id: int):
        status = self.get_collision_status(machine_id, group_id)
        if status["exhausted"]:
            raise ValueError(
                "Spazio codici LIV2 esaurito: i contatori Parte e Sottogruppo "
                "si sono sovrapposti!"
            )
        if status["warning"]:
            machine = self.get_machine(machine_id)
            group   = self.get_group(group_id)
            warnings.warn(
                f"Attenzione: rimangono solo {status['gap']} codici LIV2 "
                f"disponibili per {machine['code']}_{group['code']}",
                stacklevel=3,
            )

    def _validate_code(self, code: str, code_type: str, code_length: int):
        ok, msg = self.validate_code_string(code, code_type, code_length)
        if not ok:
            raise ValueError(f"Codice non valido '{code}': {msg}")

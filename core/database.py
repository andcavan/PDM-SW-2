# =============================================================================
#  core/database.py  –  Gestione database SQLite condiviso
# =============================================================================
import sqlite3
import threading
import time
from contextlib import contextmanager
from pathlib import Path
from typing import Optional

from filelock import FileLock, Timeout


class Database:
    """
    Wrapper attorno a SQLite su cartella di rete.
    Usa FileLock per garantire l'accesso esclusivo in scrittura.
    """

    def __init__(self, db_path: Path, lock_path: Path, timeout: int = 10):
        self.db_path   = db_path
        self.lock_path = lock_path
        self.timeout   = timeout
        self._local    = threading.local()

    # ------------------------------------------------------------------
    # Connessione
    # ------------------------------------------------------------------
    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(str(self.db_path), timeout=30)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA foreign_keys=ON")
        conn.execute("PRAGMA busy_timeout=10000")
        return conn

    @contextmanager
    def connection(self):
        """Context manager per connessione thread-locale."""
        if not hasattr(self._local, "conn") or self._local.conn is None:
            self._local.conn = self._connect()
        try:
            yield self._local.conn
        except Exception:
            self._local.conn.rollback()
            raise

    @contextmanager
    def write_lock(self):
        """Acquisisce il file lock prima di operazioni di scrittura."""
        lock = FileLock(str(self.lock_path), timeout=self.timeout)
        try:
            with lock:
                yield
        except Timeout:
            raise RuntimeError(
                "Impossibile acquisire il lock sul database. "
                "Un altro utente sta scrivendo. Riprovare tra qualche secondo."
            )

    # ------------------------------------------------------------------
    # Inizializzazione schema
    # ------------------------------------------------------------------
    def initialize(self):
        """Crea le tabelle se non esistono e applica migrazioni."""
        with self.write_lock():
            with self.connection() as conn:
                conn.executescript("""
-- UTENTI
CREATE TABLE IF NOT EXISTS users (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    username    TEXT    NOT NULL UNIQUE,
    full_name   TEXT    NOT NULL,
    role        TEXT    NOT NULL DEFAULT 'Progettista',
    workstation TEXT,
    active      INTEGER NOT NULL DEFAULT 1,
    created_at  TEXT    NOT NULL DEFAULT (datetime('now')),
    password_hash TEXT
);

-- CONFIGURAZIONE CODIFICA
CREATE TABLE IF NOT EXISTS coding_config (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    doc_type    TEXT    NOT NULL UNIQUE,   -- 'Parte','Assieme','Disegno'
    prefix      TEXT    NOT NULL DEFAULT '',
    separator   TEXT    NOT NULL DEFAULT '-',
    digits      INTEGER NOT NULL DEFAULT 5,
    start_from  INTEGER NOT NULL DEFAULT 1,
    last_number INTEGER NOT NULL DEFAULT 0,
    custom_mask TEXT                        -- es. 'PP-{NUM}-{TYPE}'
);

-- DOCUMENTI
CREATE TABLE IF NOT EXISTS documents (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    code         TEXT    NOT NULL,
    revision     TEXT    NOT NULL DEFAULT 'A',
    doc_type     TEXT    NOT NULL,          -- Parte / Assieme / Disegno
    title        TEXT    NOT NULL,
    description  TEXT,
    state        TEXT    NOT NULL DEFAULT 'In Lavorazione',
    file_name    TEXT,                      -- nome file originale
    file_ext     TEXT,                      -- .SLDPRT / .SLDASM / .SLDDRW
    archive_path TEXT,                      -- percorso relativo in archivio
    thumbnail    TEXT,
    created_by   INTEGER REFERENCES users(id),
    created_at   TEXT    NOT NULL DEFAULT (datetime('now')),
    modified_by  INTEGER REFERENCES users(id),
    modified_at  TEXT    DEFAULT (datetime('now')),
    is_locked    INTEGER NOT NULL DEFAULT 0,
    locked_by    INTEGER REFERENCES users(id),
    locked_at    TEXT,
    locked_ws    TEXT,                      -- workstation del lock
    UNIQUE(code, revision, doc_type)
);

-- VERSIONI / REVISIONI
CREATE TABLE IF NOT EXISTS document_versions (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    document_id INTEGER NOT NULL REFERENCES documents(id),
    revision    TEXT    NOT NULL,
    file_path   TEXT,
    created_by  INTEGER REFERENCES users(id),
    created_at  TEXT    NOT NULL DEFAULT (datetime('now')),
    notes       TEXT
);

-- PROPRIETÀ SOLIDWORKS
CREATE TABLE IF NOT EXISTS document_properties (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    document_id INTEGER NOT NULL REFERENCES documents(id),
    prop_name   TEXT    NOT NULL,
    prop_value  TEXT,
    UNIQUE(document_id, prop_name)
);

-- COMPONENTI ASSIEME (BOM)
CREATE TABLE IF NOT EXISTS asm_components (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    parent_id   INTEGER NOT NULL REFERENCES documents(id),
    child_id    INTEGER NOT NULL REFERENCES documents(id),
    quantity    REAL    NOT NULL DEFAULT 1,
    position    TEXT,
    notes       TEXT,
    UNIQUE(parent_id, child_id)
);

-- LOG CHECKOUT / CHECKIN
CREATE TABLE IF NOT EXISTS checkout_log (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    document_id     INTEGER NOT NULL REFERENCES documents(id),
    user_id         INTEGER NOT NULL REFERENCES users(id),
    action          TEXT    NOT NULL,       -- 'checkout' / 'checkin' / 'undo_checkout' / 'consultation'
    workstation     TEXT,
    workspace_path  TEXT,
    timestamp       TEXT    NOT NULL DEFAULT (datetime('now')),
    notes           TEXT,
    -- Snapshot al momento del checkout (per rilevare modifiche / conflitti)
    checkout_md5    TEXT,                   -- MD5 del file al momento del checkout
    checkout_size   INTEGER,               -- dimensione file al checkout
    checkout_mtime  REAL                    -- mtime file al checkout
);

-- FILE IN WORKSPACE (tracciamento copie locali)
CREATE TABLE IF NOT EXISTS workspace_files (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    document_id  INTEGER NOT NULL REFERENCES documents(id),
    user_id      INTEGER NOT NULL REFERENCES users(id),
    role         TEXT    NOT NULL DEFAULT 'checkout',  -- 'checkout' / 'component' / 'consultation'
    workspace_path TEXT,
    copied_at    TEXT    NOT NULL DEFAULT (datetime('now')),
    parent_checkout_id INTEGER,            -- se componente, ID del checkout ASM padre
    UNIQUE(document_id, user_id)
);

-- STORICO WORKFLOW
CREATE TABLE IF NOT EXISTS workflow_history (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    document_id INTEGER NOT NULL REFERENCES documents(id),
    from_state  TEXT,
    to_state    TEXT    NOT NULL,
    changed_by  INTEGER REFERENCES users(id),
    changed_at  TEXT    NOT NULL DEFAULT (datetime('now')),
    notes       TEXT
);

-- IMPOSTAZIONI CONDIVISE
CREATE TABLE IF NOT EXISTS shared_settings (
    key   TEXT PRIMARY KEY,
    value TEXT
);

-- ============================================================
-- CODIFICA GERARCHICA
-- ============================================================

-- MACCHINE (LIV0)
CREATE TABLE IF NOT EXISTS machines (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    code        TEXT    NOT NULL UNIQUE,        -- es. 'ABC', '001'
    description TEXT    NOT NULL,
    code_type   TEXT    NOT NULL DEFAULT 'ALPHA', -- 'ALPHA' | 'NUM'
    code_length INTEGER NOT NULL DEFAULT 3,
    active      INTEGER NOT NULL DEFAULT 1,
    created_at  TEXT    NOT NULL DEFAULT (datetime('now'))
);

-- GRUPPI (LIV1, legati a una macchina)
CREATE TABLE IF NOT EXISTS machine_groups (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    machine_id  INTEGER NOT NULL REFERENCES machines(id),
    code        TEXT    NOT NULL,              -- es. 'COMP', '001'
    description TEXT    NOT NULL,
    code_type   TEXT    NOT NULL DEFAULT 'ALPHA', -- 'ALPHA' | 'NUM'
    code_length INTEGER NOT NULL DEFAULT 4,
    active      INTEGER NOT NULL DEFAULT 1,
    created_at  TEXT    NOT NULL DEFAULT (datetime('now')),
    UNIQUE(machine_id, code)
);

-- CONTATORI GERARCHICI
--   counter_type = 'VERSION'  → versione LIV0 (group_id NULL)
--                               versione LIV1 (group_id valorizzato)
--   counter_type = 'PART'     → parti LIV2 (sale da 0001)
--   counter_type = 'SUBGROUP' → sottogruppi LIV2 (scende da 9999)
CREATE TABLE IF NOT EXISTS hierarchical_counters (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    counter_type TEXT    NOT NULL,
    machine_id   INTEGER REFERENCES machines(id),
    group_id     INTEGER REFERENCES machine_groups(id),
    last_value   INTEGER NOT NULL DEFAULT 0,
    UNIQUE(counter_type, machine_id, group_id)
);

-- INDICI
CREATE INDEX IF NOT EXISTS idx_doc_code      ON documents(code);
CREATE INDEX IF NOT EXISTS idx_doc_state     ON documents(state);
CREATE INDEX IF NOT EXISTS idx_doc_type      ON documents(doc_type);
CREATE INDEX IF NOT EXISTS idx_doc_locked    ON documents(is_locked);
CREATE INDEX IF NOT EXISTS idx_asm_parent    ON asm_components(parent_id);
CREATE INDEX IF NOT EXISTS idx_asm_child     ON asm_components(child_id);
CREATE INDEX IF NOT EXISTS idx_mgrp_machine  ON machine_groups(machine_id);
CREATE INDEX IF NOT EXISTS idx_hcnt_type     ON hierarchical_counters(counter_type, machine_id, group_id);
                """)
                conn.commit()
        # Ripara eventuali FK corrotti da migration precedente
        self._repair_stale_fk_references()
        # Esegui migrazioni colonne aggiuntive su tabelle esistenti
        self._migrate()
        self._migrate_documents_unique()

    def _migrate_documents_unique(self):
        """
        Aggiorna il constraint UNIQUE su documents da (code, revision)
        a (code, revision, doc_type).

        Gestisce anche il recovery da stato rotto (crash a metà migration):
        - Se esiste _documents_old ma non documents → completa la migration
        - Se esiste _documents_old e esiste documents → elimina il residuo
        - Se documents ha già il nuovo constraint → non fa nulla
        """
        NEW_TABLE_DDL = """
            CREATE TABLE documents (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                code         TEXT    NOT NULL,
                revision     TEXT    NOT NULL DEFAULT '00',
                doc_type     TEXT    NOT NULL,
                title        TEXT    NOT NULL,
                description  TEXT,
                state        TEXT    NOT NULL DEFAULT 'In Lavorazione',
                file_name    TEXT,
                file_ext     TEXT,
                archive_path TEXT,
                thumbnail    TEXT,
                created_by   INTEGER REFERENCES users(id),
                created_at   TEXT    NOT NULL DEFAULT (datetime('now')),
                modified_by  INTEGER REFERENCES users(id),
                modified_at  TEXT    DEFAULT (datetime('now')),
                is_locked    INTEGER NOT NULL DEFAULT 0,
                locked_by    INTEGER REFERENCES users(id),
                locked_at    TEXT,
                locked_ws    TEXT,
                machine_id   INTEGER REFERENCES machines(id),
                group_id     INTEGER REFERENCES machine_groups(id),
                doc_level    INTEGER DEFAULT 2,
                parent_doc_id INTEGER REFERENCES documents(id),
                checkout_md5    TEXT,
                checkout_size   INTEGER,
                checkout_mtime  REAL,
                UNIQUE(code, revision, doc_type)
            )
        """

        # Colonne da copiare nella migration (deve corrispondere a NEW_TABLE_DDL)
        _COPY_COLS = """id, code, revision, doc_type, title, description,
                       state, file_name, file_ext, archive_path, thumbnail,
                       created_by, created_at, modified_by, modified_at,
                       is_locked, locked_by, locked_at, locked_ws,
                       machine_id, group_id, doc_level, parent_doc_id"""

        with self.write_lock():
            with self.connection() as conn:
                conn.execute("PRAGMA foreign_keys=OFF")

                # Stato delle tabelle
                tables = {
                    r[0] for r in conn.execute(
                        "SELECT name FROM sqlite_master WHERE type='table'"
                    ).fetchall()
                }
                has_docs     = "documents"      in tables
                has_old      = "_documents_old" in tables

                # ── Recovery: crash a metà migration precedente ──────────
                if has_old and not has_docs:
                    # _documents_old esiste ma documents no → completiamo
                    conn.execute(NEW_TABLE_DDL)
                    # Copia solo le colonne presenti nella vecchia tabella
                    old_cols = {r[1] for r in conn.execute(
                        "PRAGMA table_info(_documents_old)").fetchall()}
                    copy_cols = [c.strip() for c in _COPY_COLS.split(",")
                                 if c.strip() in old_cols]
                    cols_str = ", ".join(copy_cols)
                    conn.execute(f"""
                        INSERT INTO documents ({cols_str})
                        SELECT {cols_str} FROM _documents_old
                    """)
                    conn.execute("DROP TABLE _documents_old")
                    self._create_doc_indexes(conn)
                    conn.execute("PRAGMA foreign_keys=ON")
                    conn.commit()
                    return

                if has_old and has_docs:
                    # Residuo da migration già completata → pulizia
                    conn.execute("DROP TABLE _documents_old")
                    conn.commit()

                if not has_docs:
                    conn.execute("PRAGMA foreign_keys=ON")
                    conn.commit()
                    return

                # ── Controllo se migration è già stata applicata ──────────
                schema_row = conn.execute(
                    "SELECT sql FROM sqlite_master "
                    "WHERE type='table' AND name='documents'"
                ).fetchone()
                schema_sql = (schema_row[0] if isinstance(schema_row, tuple)
                              else (schema_row["sql"] if schema_row else "")) or ""
                if "code, revision, doc_type" in schema_sql:
                    conn.execute("PRAGMA foreign_keys=ON")
                    conn.commit()
                    return

                # ── Migration: replace constraint ─────────────────────────
                # IMPORTANTE: legacy_alter_table=ON evita che SQLite >= 3.25
                # aggiorni i FK di ALTRE tabelle (checkout_log, ecc.)
                # facendoli puntare a '_documents_old' invece di 'documents'
                try:
                    conn.execute("PRAGMA legacy_alter_table=ON")
                    # Colonne presenti nella vecchia tabella
                    old_cols = {r[1] for r in conn.execute(
                        "PRAGMA table_info(documents)").fetchall()}
                    conn.execute(
                        "ALTER TABLE documents RENAME TO _documents_old"
                    )
                    conn.execute(NEW_TABLE_DDL)
                    copy_cols = [c.strip() for c in _COPY_COLS.split(",")
                                 if c.strip() in old_cols]
                    cols_str = ", ".join(copy_cols)
                    conn.execute(f"""
                        INSERT INTO documents ({cols_str})
                        SELECT {cols_str} FROM _documents_old
                    """)
                    conn.execute("DROP TABLE _documents_old")
                    self._create_doc_indexes(conn)
                    conn.execute("PRAGMA legacy_alter_table=OFF")
                    conn.execute("PRAGMA foreign_keys=ON")
                    conn.commit()
                except Exception:
                    conn.rollback()
                    conn.execute("PRAGMA legacy_alter_table=OFF")
                    conn.execute("PRAGMA foreign_keys=ON")
                    raise

    def _create_doc_indexes(self, conn):
        """Ricrea gli indici sulla tabella documents."""
        for idx_sql in [
            "CREATE INDEX IF NOT EXISTS idx_doc_code   ON documents(code)",
            "CREATE INDEX IF NOT EXISTS idx_doc_state  ON documents(state)",
            "CREATE INDEX IF NOT EXISTS idx_doc_type   ON documents(doc_type)",
            "CREATE INDEX IF NOT EXISTS idx_doc_locked ON documents(is_locked)",
        ]:
            conn.execute(idx_sql)

    def _repair_stale_fk_references(self):
        """
        Ripara i FK corrotti da migration precedente.

        Bug SQLite >= 3.25: ALTER TABLE RENAME aggiorna automaticamente
        i riferimenti FK in TUTTE le tabelle. Se 'documents' viene rinominata
        in '_documents_old', checkout_log e altre tabelle avranno i FK che
        puntano a '_documents_old' (che poi viene droppata).

        Fix: usa PRAGMA writable_schema per correggere direttamente sqlite_master.
        """
        with self.write_lock():
            with self.connection() as conn:
                stale = conn.execute(
                    "SELECT name, sql FROM sqlite_master "
                    "WHERE sql LIKE '%_documents_old%' AND type='table'"
                ).fetchall()
                if not stale:
                    return
                conn.execute("PRAGMA writable_schema=ON")
                for row in stale:
                    name = row[0] if isinstance(row, tuple) else row["name"]
                    old_sql = row[1] if isinstance(row, tuple) else row["sql"]
                    new_sql = old_sql.replace("_documents_old", "documents")
                    conn.execute(
                        "UPDATE sqlite_master SET sql=? WHERE type='table' AND name=?",
                        (new_sql, name),
                    )
                conn.execute("PRAGMA writable_schema=OFF")
                # Forza SQLite a ricaricare lo schema
                conn.execute("PRAGMA integrity_check")
                conn.commit()

    # ------------------------------------------------------------------
    # Migrazione colonne su tabelle esistenti
    # ------------------------------------------------------------------
    def _migrate(self):
        """Aggiunge colonne mancanti alla tabella documents (migrazioni)."""
        new_cols = [
            ("machine_id",       "INTEGER REFERENCES machines(id)"),
            ("group_id",         "INTEGER REFERENCES machine_groups(id)"),
            ("doc_level",        "INTEGER DEFAULT 2"),
            ("parent_doc_id",    "INTEGER REFERENCES documents(id)"),
            # Snapshot checkout per rilevamento modifiche / conflitti
            ("checkout_md5",     "TEXT"),
            ("checkout_size",    "INTEGER"),
            ("checkout_mtime",   "REAL"),
        ]
        # Migrazioni checkout_log
        checkout_log_cols = [
            ("checkout_md5",   "TEXT"),
            ("checkout_size",  "INTEGER"),
            ("checkout_mtime", "REAL"),
        ]
        with self.connection() as conn:
            cur = conn.execute("PRAGMA table_info(documents)")
            existing = {row[1] for row in cur.fetchall()}
            cur2 = conn.execute("PRAGMA table_info(checkout_log)")
            existing_cl = {row[1] for row in cur2.fetchall()}
        for col_name, col_def in new_cols:
            if col_name not in existing:
                try:
                    with self.write_lock():
                        with self.connection() as conn:
                            conn.execute(
                                f"ALTER TABLE documents ADD COLUMN {col_name} {col_def}"
                            )
                            conn.commit()
                except Exception:
                    pass
        for col_name, col_def in checkout_log_cols:
            if col_name not in existing_cl:
                try:
                    with self.write_lock():
                        with self.connection() as conn:
                            conn.execute(
                                f"ALTER TABLE checkout_log ADD COLUMN {col_name} {col_def}"
                            )
                            conn.commit()
                except Exception:
                    pass
        # Tabella workspace_files (creazione se non esiste)
        try:
            with self.write_lock():
                with self.connection() as conn:
                    conn.execute("""
                        CREATE TABLE IF NOT EXISTS workspace_files (
                            id           INTEGER PRIMARY KEY AUTOINCREMENT,
                            document_id  INTEGER NOT NULL REFERENCES documents(id),
                            user_id      INTEGER NOT NULL REFERENCES users(id),
                            role         TEXT    NOT NULL DEFAULT 'checkout',
                            workspace_path TEXT,
                            copied_at    TEXT    NOT NULL DEFAULT (datetime('now')),
                            parent_checkout_id INTEGER,
                            UNIQUE(document_id, user_id)
                        )
                    """)
                    conn.commit()
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Helper generici
    # ------------------------------------------------------------------
    def fetchall(self, sql: str, params=()) -> list:
        with self.connection() as conn:
            cur = conn.execute(sql, params)
            return [dict(r) for r in cur.fetchall()]

    def fetchone(self, sql: str, params=()) -> Optional[dict]:
        with self.connection() as conn:
            cur = conn.execute(sql, params)
            row = cur.fetchone()
            return dict(row) if row else None

    def execute(self, sql: str, params=()) -> int:
        """Esegue una scrittura e ritorna lastrowid."""
        with self.write_lock():
            with self.connection() as conn:
                cur = conn.execute(sql, params)
                conn.commit()
                return cur.lastrowid

    def executemany(self, sql: str, params_list: list):
        with self.write_lock():
            with self.connection() as conn:
                conn.executemany(sql, params_list)
                conn.commit()

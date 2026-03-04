# =============================================================================
#  ui/session.py  –  Sessione applicazione (singleton globale)
# =============================================================================
from pathlib import Path
from typing import Optional

from config import SharedPaths, load_local_config
from core.database import Database
from core.user_manager import UserManager
from core.coding_manager import CodingManager
from core.file_manager import FileManager
from core.checkout_manager import CheckoutManager
from core.workflow_manager import WorkflowManager
from core.properties_manager import PropertiesManager
from core.asm_manager import AsmManager


class AppSession:
    """Contenitore della sessione corrente: utente, DB e manager."""

    def __init__(self):
        cfg             = load_local_config()
        shared_root     = cfg.get("shared_root", "")
        self.sp         = SharedPaths(shared_root) if shared_root else None
        self.db: Optional[Database] = None
        self.user: Optional[dict]   = None

        # Manager (inizializzati dopo connessione DB)
        self.users:      Optional[UserManager]       = None
        self.coding:     Optional[CodingManager]     = None
        self.files:      Optional[FileManager]       = None
        self.checkout:   Optional[CheckoutManager]   = None
        self.workflow:   Optional[WorkflowManager]   = None
        self.properties: Optional[PropertiesManager] = None
        self.asm:        Optional[AsmManager]        = None

    # ------------------------------------------------------------------
    def connect(self, shared_root: str):
        """Connette al database sulla cartella condivisa."""
        self.sp = SharedPaths(shared_root)
        self.sp.ensure_dirs()
        self.db = Database(self.sp.db_file, self.sp.db_lock_file)
        self.db.initialize()

        self.users      = UserManager(self.db)
        self.coding     = CodingManager(self.db)
        self.workflow   = WorkflowManager(self.db)
        self.properties = PropertiesManager(self.db)
        self.asm        = AsmManager(self.db)

    def set_user(self, user: dict):
        self.user = user
        if self.db and self.sp:
            self.files    = FileManager(self.db, self.sp, user)
            self.checkout = CheckoutManager(self.db, self.sp, user)

    @property
    def is_connected(self) -> bool:
        return self.db is not None

    @property
    def is_logged_in(self) -> bool:
        return self.user is not None

    def can(self, permission: str) -> bool:
        if not self.user or not self.users:
            return False
        return self.users.has_permission(self.user, permission)

    def logout(self):
        self.user     = None
        self.files    = None
        self.checkout = None


# Istanza globale
session = AppSession()

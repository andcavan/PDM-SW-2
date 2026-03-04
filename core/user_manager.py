# =============================================================================
#  core/user_manager.py  –  Gestione utenti
# =============================================================================
from __future__ import annotations
import hashlib
import socket
from typing import Optional, TYPE_CHECKING
if TYPE_CHECKING:
    from core.database import Database

from config import USER_ROLES, ROLE_PERMISSIONS


def _hash(password: str) -> str:
    return hashlib.sha256(password.encode()).hexdigest()


class UserManager:
    def __init__(self, db: "Database"):
        self.db = db
        self._ensure_admin()

    def _ensure_admin(self):
        """Crea l'utente amministratore predefinito se non esiste."""
        if not self.db.fetchone("SELECT id FROM users WHERE username='admin'"):
            self.db.execute(
                """INSERT INTO users (username, full_name, role, password_hash)
                   VALUES ('admin','Amministratore','Amministratore',?)""",
                (_hash("admin"),),
            )

    # ------------------------------------------------------------------
    def authenticate(self, username: str, password: str) -> Optional[dict]:
        user = self.db.fetchone(
            "SELECT * FROM users WHERE username=? AND active=1", (username,)
        )
        if not user:
            return None
        stored = user.get("password_hash") or ""
        if not stored or stored == _hash(password):
            # Aggiorna workstation
            self.db.execute(
                "UPDATE users SET workstation=? WHERE id=?",
                (socket.gethostname(), user["id"]),
            )
            return dict(user)
        return None

    def login_no_password(self, username: str) -> Optional[dict]:
        """Login senza password (modalità rete senza autenticazione)."""
        return self.db.fetchone(
            "SELECT * FROM users WHERE username=? AND active=1", (username,)
        )

    # ------------------------------------------------------------------
    def create_user(self, username: str, full_name: str, role: str,
                    password: str = "") -> int:
        if role not in USER_ROLES:
            raise ValueError(f"Ruolo non valido: {role}")
        return self.db.execute(
            """INSERT INTO users (username, full_name, role, password_hash)
               VALUES (?,?,?,?)""",
            (username, full_name, role, _hash(password) if password else ""),
        )

    def update_user(self, user_id: int, full_name: str, role: str,
                    password: str = "", active: bool = True):
        if password:
            self.db.execute(
                """UPDATE users SET full_name=?, role=?, active=?, password_hash=?
                   WHERE id=?""",
                (full_name, role, int(active), _hash(password), user_id),
            )
        else:
            self.db.execute(
                "UPDATE users SET full_name=?, role=?, active=? WHERE id=?",
                (full_name, role, int(active), user_id),
            )

    def delete_user(self, user_id: int):
        self.db.execute(
            "UPDATE users SET active=0 WHERE id=?", (user_id,)
        )

    def get_all_users(self) -> list:
        return self.db.fetchall(
            "SELECT * FROM users WHERE active=1 ORDER BY full_name"
        )

    def get_user(self, user_id: int) -> Optional[dict]:
        return self.db.fetchone("SELECT * FROM users WHERE id=?", (user_id,))

    # ------------------------------------------------------------------
    def has_permission(self, user: dict, permission: str) -> bool:
        role = user.get("role", "Utente")
        perms = ROLE_PERMISSIONS.get(role, {})
        return perms.get(permission, False)

# =============================================================================
#  PDM-SW  –  Configurazione globale
# =============================================================================
import json
import os
from pathlib import Path

APP_NAME    = "PDM-SW"
APP_VERSION = "2.1.2"

# Cartella locale dell'applicazione (stessa posizione di questo file)
APP_DIR = Path(__file__).parent.resolve()

# File di configurazione locale (percorso cartella condivisa)
LOCAL_CFG_FILE = APP_DIR / "local_config.json"


def load_local_config() -> dict:
    """Carica la configurazione locale (percorso rete, utente, ecc.)."""
    if LOCAL_CFG_FILE.exists():
        with open(LOCAL_CFG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_local_config(cfg: dict):
    """Salva la configurazione locale."""
    with open(LOCAL_CFG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)


class SharedPaths:
    """Percorsi relativi alla cartella condivisa di rete."""

    def __init__(self, shared_root: str):
        self.root       = Path(shared_root)
        self.database   = self.root / "database"
        self.archive    = self.root / "archive"
        self.workspace  = self.root / "workspace"
        self.config     = self.root / "config"
        self.thumbnails = self.root / "thumbnails"
        self.temp       = self.root / "temp"

    def ensure_dirs(self):
        for d in [self.database, self.archive, self.workspace,
                  self.config, self.thumbnails, self.temp]:
            d.mkdir(parents=True, exist_ok=True)

    @property
    def db_file(self) -> Path:
        return self.database / "pdm.db"

    @property
    def db_lock_file(self) -> Path:
        return self.database / "pdm.lock"

    @property
    def shared_settings_file(self) -> Path:
        return self.config / "shared_settings.json"

    def archive_path(self, code: str, revision: str) -> Path:
        return self.archive / code / revision

    def workspace_path(self, username: str) -> Path:
        return self.workspace / username


# Colori workflow (stato documento)
WORKFLOW_STATES = {
    "In Lavorazione": {"color": "#2196F3", "icon": "work"},
    "In Revisione":   {"color": "#FF9800", "icon": "review"},
    "Revisionato":    {"color": "#9C27B0", "icon": "revised"},
    "Rilasciato":     {"color": "#4CAF50", "icon": "released"},
    "Obsoleto":       {"color": "#9E9E9E", "icon": "obsolete"},
}

# Transizioni workflow consentite
WORKFLOW_TRANSITIONS = {
    "In Lavorazione": ["In Revisione", "Rilasciato"],
    "In Revisione":   ["Revisionato", "In Lavorazione"],
    "Revisionato":    ["Rilasciato", "In Lavorazione"],
    "Rilasciato":     ["In Revisione", "Obsoleto"],
    "Obsoleto":       [],
}

# Tipi di documento SolidWorks
SW_EXTENSIONS = {
    ".SLDPRT": "Parte",
    ".SLDASM": "Assieme",
    ".SLDDRW": "Disegno",
}

# Mapping: estensione template SolidWorks → estensione file finale
TPL_EXT_MAP = {
    ".prtdot": ".SLDPRT",
    ".asmdot": ".SLDASM",
    ".drwdot": ".SLDDRW",
    ".PRTDOT": ".SLDPRT",
    ".ASMDOT": ".SLDASM",
    ".DRWDOT": ".SLDDRW",
}

# Ruoli utente
USER_ROLES = ["Utente", "Progettista", "Responsabile", "Amministratore"]

# Permessi per ruolo
ROLE_PERMISSIONS = {
    "Utente": {
        "checkout": False, "checkin": False, "create": False,
        "release": False, "admin": False,
    },
    "Progettista": {
        "checkout": True, "checkin": True, "create": True,
        "release": False, "admin": False,
    },
    "Responsabile": {
        "checkout": True, "checkin": True, "create": True,
        "release": True, "admin": False,
    },
    "Amministratore": {
        "checkout": True, "checkin": True, "create": True,
        "release": True, "admin": True,
    },
}

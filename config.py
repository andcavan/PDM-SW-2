# =============================================================================
#  PDM-SW  –  Configurazione globale
# =============================================================================
import copy
import json
import os
from pathlib import Path

APP_NAME    = "PDM-SW"
APP_VERSION = "2.10.19"

# Cartella locale dell'applicazione (stessa posizione di questo file)
APP_DIR = Path(__file__).parent.resolve()

# File puntatore alla cartella dati locali (creato al primo avvio)
DATA_DIR_FILE = APP_DIR / ".pdm_datadir"


def get_data_dir() -> Path:
    """Restituisce la cartella dati locali (dove risiede local_config.json).
    Legge .pdm_datadir se esiste, altrimenti usa APP_DIR come fallback."""
    if DATA_DIR_FILE.exists():
        path_str = DATA_DIR_FILE.read_text(encoding="utf-8").strip()
        if path_str:
            return Path(path_str)
    return APP_DIR


def _get_local_cfg_file() -> Path:
    return get_data_dir() / "local_config.json"

# Chiavi che appartengono ai profili (non globali)
PROFILE_KEYS = frozenset({
    "shared_root",
    "sw_exe_path", "edrawings_exe_path",
    "sw_template_prt", "sw_template_asm", "sw_template_drw",
    "sw_reg_file", "sw_workspace",
    "sw_property_mapping",
})


# ---- Raw I/O (formato interno) ------------------------------------------

def _load_raw_config() -> dict:
    cfg_file = _get_local_cfg_file()
    if cfg_file.exists():
        with open(cfg_file, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def _save_raw_config(raw: dict):
    cfg_file = _get_local_cfg_file()
    cfg_file.parent.mkdir(parents=True, exist_ok=True)
    with open(cfg_file, "w", encoding="utf-8") as f:
        json.dump(raw, f, indent=2, ensure_ascii=False)


def _ensure_profile_format(raw: dict) -> dict:
    """Migra dalla config flat al formato profili (una tantum)."""
    if "profiles" in raw:
        return raw
    profile_data = {}
    global_data = {}
    for k, v in raw.items():
        if k in PROFILE_KEYS:
            profile_data[k] = v
        else:
            global_data[k] = v
    if profile_data.get("shared_root"):
        global_data["active_profile"] = "Default"
        global_data["profiles"] = {"Default": profile_data}
    else:
        global_data["active_profile"] = ""
        global_data["profiles"] = {}
    _save_raw_config(global_data)
    return global_data


# ---- API pubblica (backward-compatible) ----------------------------------

def load_local_config() -> dict:
    """Carica configurazione effettiva (globale + profilo attivo, flat)."""
    raw = _ensure_profile_format(_load_raw_config())
    profile_name = raw.get("active_profile", "")
    profile = raw.get("profiles", {}).get(profile_name, {})
    result = {k: v for k, v in raw.items() if k != "profiles"}
    result.update(profile)
    return result


def save_local_config(cfg: dict):
    """Salva la configurazione. Chiavi profilo → profilo attivo, resto → globale."""
    raw = _ensure_profile_format(_load_raw_config())
    profile_name = cfg.get("active_profile", raw.get("active_profile", ""))

    if not profile_name and cfg.get("shared_root"):
        profile_name = "Default"

    raw["active_profile"] = profile_name
    profiles = raw.setdefault("profiles", {})
    profile_data = profiles.get(profile_name, {}) if profile_name else {}

    for k, v in cfg.items():
        if k in ("profiles", "active_profile"):
            continue
        elif k in PROFILE_KEYS:
            profile_data[k] = v
        else:
            raw[k] = v

    # Migra workspace in sospeso (_init_workspace) al primo profilo creato
    if profile_name and profile_name not in profiles:
        pending_ws = raw.pop("_init_workspace", None)
        if pending_ws:
            profile_data.setdefault("sw_workspace", pending_ws)

    if profile_name:
        profiles[profile_name] = profile_data
    _save_raw_config(raw)


# ---- Gestione profili ----------------------------------------------------

def get_profile_names() -> list:
    raw = _ensure_profile_format(_load_raw_config())
    return list(raw.get("profiles", {}).keys())


def get_active_profile_name() -> str:
    raw = _ensure_profile_format(_load_raw_config())
    return raw.get("active_profile", "")


def set_active_profile(name: str):
    raw = _ensure_profile_format(_load_raw_config())
    if name not in raw.get("profiles", {}):
        raise ValueError(f"Profilo '{name}' non trovato")
    raw["active_profile"] = name
    _save_raw_config(raw)


def load_profile(name: str) -> dict:
    raw = _ensure_profile_format(_load_raw_config())
    return dict(raw.get("profiles", {}).get(name, {}))


def save_profile(name: str, data: dict):
    raw = _ensure_profile_format(_load_raw_config())
    raw.setdefault("profiles", {})[name] = data
    _save_raw_config(raw)


def delete_profile(name: str):
    raw = _ensure_profile_format(_load_raw_config())
    profiles = raw.get("profiles", {})
    if name not in profiles:
        return
    del profiles[name]
    if raw.get("active_profile") == name:
        raw["active_profile"] = next(iter(profiles), "")
    _save_raw_config(raw)


def rename_profile(old_name: str, new_name: str):
    raw = _ensure_profile_format(_load_raw_config())
    profiles = raw.get("profiles", {})
    if old_name not in profiles:
        raise ValueError(f"Profilo '{old_name}' non trovato")
    if new_name in profiles:
        raise ValueError(f"Profilo '{new_name}' già esistente")
    profiles[new_name] = profiles.pop(old_name)
    if raw.get("active_profile") == old_name:
        raw["active_profile"] = new_name
    _save_raw_config(raw)


def copy_profile(src_name: str, dst_name: str) -> dict:
    """Copia configurazione di un profilo in un nuovo profilo. Ritorna i dati copiati."""
    raw = _ensure_profile_format(_load_raw_config())
    profiles = raw.get("profiles", {})
    if src_name not in profiles:
        raise ValueError(f"Profilo '{src_name}' non trovato")
    if dst_name in profiles:
        raise ValueError(f"Profilo '{dst_name}' già esistente")
    new_data = copy.deepcopy(profiles[src_name])
    profiles[dst_name] = new_data
    _save_raw_config(raw)
    return new_data


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
    "Rilasciato":     {"color": "#4CAF50", "icon": "released"},
    "In Revisione":   {"color": "#FF9800", "icon": "review"},
    "Obsoleto":       {"color": "#9E9E9E", "icon": "obsolete"},
}

# Transizioni workflow consentite (4 stati)
# In Lavorazione → Rilasciato   (prima emissione)
# In Revisione   → Rilasciato   (chiude revisione, incrementa rev)
# Rilasciato     → Obsoleto     (fuori produzione manuale)
# Obsoleto       → (nessuna)    stato finale
# Nota: «Crea revisione» da Rilasciato è un'operazione separata
WORKFLOW_TRANSITIONS = {
    "In Lavorazione": ["Rilasciato"],
    "In Revisione":   ["Rilasciato"],
    "Rilasciato":     ["Obsoleto"],
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

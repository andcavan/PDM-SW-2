# =============================================================================
#  core/coding_config.py  –  Configurazione schema di codifica
#
#  Permette di personalizzare il formato dei codici per ogni livello gerarchico
#  senza modificare il codice sorgente. La configurazione viene salvata nel DB
#  (shared_settings, key='coding_scheme_config') come JSON.
#
#  Variabili template supportate:
#    {MACH}    → codice macchina (es. ABC)
#    {GRP}     → codice gruppo (es. COMP)
#    {VER:N}   → numero versione, N cifre con zfill  (es. {VER:3} → 001)
#    {NUM:N}   → numero sequenziale (sale o scende), N cifre (es. {NUM:4} → 0001)
#
#  Prefisso e suffisso vengono concatenati al codice prodotto dal template:
#    codice_finale = prefix + render(template) + suffix
# =============================================================================
from __future__ import annotations
import json
import re
from dataclasses import dataclass, field, asdict


@dataclass
class LevelConfig:
    """Configurazione per un singolo livello di codifica."""
    template: str            # template con variabili {MACH}, {GRP}, {VER:N}, {NUM:N}
    prefix: str  = ""        # prefisso fisso aggiunto prima del codice generato
    suffix: str  = ""        # suffisso fisso aggiunto dopo il codice generato
    num_start: int   = 1     # valore di partenza del contatore
    num_max: int     = 9999  # limite superiore (usato se num_dir="asc")
    num_min: int     = 1     # limite inferiore (usato se num_dir="desc")
    num_dir: str = "asc"     # "asc" = sale da num_start, "desc" = scende da num_start
    collision_threshold: int = 500  # warning quando rimangono meno di N codici


@dataclass
class CodingSchemeConfig:
    """Configurazione completa dello schema di codifica a 4 livelli."""

    name: str = "Standard Gerarchico"

    # LIV0 – Macchina (ASM): es. ABC_V001
    liv0: LevelConfig = field(default_factory=lambda: LevelConfig(
        template="{MACH}_V{VER:3}",
        num_start=1, num_max=9999, num_dir="asc"
    ))

    # LIV1 – Gruppo (ASM): es. ABC_COMP-V001
    liv1: LevelConfig = field(default_factory=lambda: LevelConfig(
        template="{MACH}_{GRP}-V{VER:3}",
        num_start=1, num_max=9999, num_dir="asc"
    ))

    # LIV2/1 – Sottogruppo (ASM): numeri discendenti, es. ABC_COMP-9999
    liv2_1: LevelConfig = field(default_factory=lambda: LevelConfig(
        template="{MACH}_{GRP}-{NUM:4}",
        num_start=9999, num_min=9000, num_dir="desc",
        collision_threshold=500
    ))

    # LIV2/2 – Parte (PRT): numeri ascendenti, es. ABC_COMP-0001
    liv2_2: LevelConfig = field(default_factory=lambda: LevelConfig(
        template="{MACH}_{GRP}-{NUM:4}",
        num_start=1, num_max=8999, num_dir="asc"
    ))

    # Formato codici macchina e gruppo (definiti a livello schema)
    mach_code_type: str   = "ALPHA"  # "ALPHA", "NUM" o "ALPHA+NUM"
    mach_code_length: int = 3
    grp_code_type: str    = "ALPHA"
    grp_code_length: int  = 4

    # ------------------------------------------------------------------
    # Serializzazione / deserializzazione
    # ------------------------------------------------------------------

    def to_json(self) -> str:
        """Serializza la config in JSON."""
        return json.dumps(asdict(self), indent=2, ensure_ascii=False)

    @classmethod
    def from_json(cls, text: str) -> "CodingSchemeConfig":
        """Deserializza la config da JSON. Ritorna il default in caso di errore."""
        data = json.loads(text)
        cfg = cls()
        cfg.name = data.get("name", cfg.name)
        for attr in ("liv0", "liv1", "liv2_1", "liv2_2"):
            raw = data.get(attr)
            if raw and isinstance(raw, dict):
                try:
                    setattr(cfg, attr, LevelConfig(**raw))
                except TypeError:
                    pass  # campo sconosciuto o mancante → mantieni default
        # Campi semplici
        for attr in ("mach_code_type", "mach_code_length",
                     "grp_code_type", "grp_code_length"):
            if attr in data:
                setattr(cfg, attr, data[attr])
        return cfg

    @classmethod
    def default(cls) -> "CodingSchemeConfig":
        """Ritorna una configurazione con i valori di default (=comportamento storico)."""
        return cls()

    # ------------------------------------------------------------------
    # Rendering template
    # ------------------------------------------------------------------

    @staticmethod
    def render_template(template: str, mach: str = "", grp: str = "",
                        num: int = 0, ver: int = 0) -> str:
        """
        Espande le variabili del template e ritorna il codice risultante.

        Sostituzioni:
          {MACH}    → mach (stringa, non modificata)
          {GRP}     → grp  (stringa, non modificata)
          {VER:N}   → str(ver).zfill(N)
          {NUM:N}   → str(num).zfill(N)
        """
        result = template

        # {MACH} e {GRP}
        result = result.replace("{MACH}", mach)
        result = result.replace("{GRP}",  grp)

        # {VER:N}
        for m in re.findall(r'\{VER:(\d+)\}', result):
            digits = int(m)
            result = result.replace(f"{{VER:{m}}}", str(ver).zfill(digits))

        # {NUM:N}
        for m in re.findall(r'\{NUM:(\d+)\}', result):
            digits = int(m)
            result = result.replace(f"{{NUM:{m}}}", str(num).zfill(digits))

        return result

    def render(self, level_cfg: LevelConfig, mach: str = "", grp: str = "",
               num: int = 0, ver: int = 0) -> str:
        """
        Compone il codice finale: prefix + render_template(...) + suffix.
        """
        body = self.render_template(level_cfg.template, mach=mach, grp=grp,
                                    num=num, ver=ver)
        return level_cfg.prefix + body + level_cfg.suffix

    # ------------------------------------------------------------------
    # Validazione template
    # ------------------------------------------------------------------

    KNOWN_VARS = re.compile(
        r'\{MACH\}|\{GRP\}|\{VER:\d+\}|\{NUM:\d+\}'
    )
    BRACE_VARS = re.compile(r'\{[^}]+\}')

    @staticmethod
    def validate_template(template: str) -> list[str]:
        """
        Controlla il template e ritorna lista di variabili non riconosciute.
        Lista vuota = template valido.
        """
        unknown = []
        for m in re.findall(r'\{[^}]+\}', template):
            if not re.fullmatch(
                r'\{MACH\}|\{GRP\}|\{VER:\d+\}|\{NUM:\d+\}', m
            ):
                unknown.append(m)
        return unknown

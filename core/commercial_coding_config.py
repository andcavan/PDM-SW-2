# =============================================================================
#  core/commercial_coding_config.py  –  Configurazione schema codifica commerciali
#
#  Gestisce il formato dei codici per articoli commerciali e normalizzati.
#  La configurazione viene salvata nel DB (shared_settings, chiave
#  'commercial_coding_config') come JSON.
#
#  Formato codice:
#    {L}-{CAT}-{SUB}-{NUM:4}   (con sottocategoria)
#    {L}-{CAT}-{NUM:4}         (senza sottocategoria)
#
#  Variabili template:
#    {L}      → prefisso tipo: "5" (commerciale) o "6" (normalizzato)
#    {CAT}    → codice categoria (es. VIT, CUS, ELE)
#    {SUB}    → codice sottocategoria (es. ISO, DIN, SKF)
#    {NUM:N}  → numero sequenziale, N cifre con zfill (es. {NUM:4} → 0001)
# =============================================================================
from __future__ import annotations
import json
import re
from dataclasses import dataclass, asdict


@dataclass
class CommercialCodingConfig:
    """Configurazione schema codifica articoli commerciali/normalizzati."""

    name: str = "Standard Commerciali"

    # Template con sottocategoria: {L}-{CAT}-{SUB}-{NUM:4}
    template_with_sub: str = "{L}-{CAT}-{SUB}-{NUM:4}"

    # Template senza sottocategoria: {L}-{CAT}-{NUM:4}
    template_without_sub: str = "{L}-{CAT}-{NUM:4}"

    # Lunghezza numero sequenziale (cifre con zfill)
    num_digits: int = 4

    # Valore di partenza contatore
    num_start: int = 1

    # Limite massimo contatore
    num_max: int = 9999

    # Lunghezza codice categoria (4 cifre numeriche auto-incrementali)
    cat_code_length: int = 4

    # Lunghezza codice sottocategoria (4 cifre numeriche auto-incrementali)
    sub_code_length: int = 4

    # Percorso cartella archivio file SW commerciali.
    # Vuoto = usa default {SharedPaths.archive}/commercial/
    # Relativo = relativo a SharedPaths.root
    # Assoluto = usato direttamente
    commercial_archive_path: str = ""

    # ------------------------------------------------------------------
    # Serializzazione / deserializzazione
    # ------------------------------------------------------------------

    def to_json(self) -> str:
        return json.dumps(asdict(self), indent=2, ensure_ascii=False)

    @classmethod
    def from_json(cls, text: str) -> "CommercialCodingConfig":
        """Deserializza da JSON. Ritorna il default in caso di errore."""
        try:
            data = json.loads(text)
        except (json.JSONDecodeError, TypeError):
            return cls()
        cfg = cls()
        for attr in (
            "name", "template_with_sub", "template_without_sub",
            "num_digits", "num_start", "num_max",
            "cat_code_length", "sub_code_length",
            "commercial_archive_path",
        ):
            if attr in data:
                setattr(cfg, attr, data[attr])
        return cfg

    @classmethod
    def default(cls) -> "CommercialCodingConfig":
        return cls()

    # ------------------------------------------------------------------
    # Rendering
    # ------------------------------------------------------------------

    @staticmethod
    def render_template(template: str, prefix_l: str = "",
                        cat: str = "", sub: str = "", num: int = 0) -> str:
        """
        Espande le variabili del template e ritorna il codice risultante.

        Sostituzioni:
          {L}      → prefix_l (es. "5" o "6")
          {CAT}    → cat (es. "VIT")
          {SUB}    → sub (es. "ISO")
          {NUM:N}  → str(num).zfill(N)
        """
        result = template
        result = result.replace("{L}",   prefix_l)
        result = result.replace("{CAT}", cat)
        result = result.replace("{SUB}", sub)
        for m in re.findall(r'\{NUM:(\d+)\}', result):
            digits = int(m)
            result = result.replace(f"{{NUM:{m}}}", str(num).zfill(digits))
        return result

    def render(self, prefix_l: str, cat: str, sub: str | None,
               num: int) -> str:
        """
        Seleziona il template corretto e produce il codice finale.

        Args:
            prefix_l: "5" o "6"
            cat:      codice categoria
            sub:      codice sottocategoria (None o "" per template senza SUB)
            num:      valore contatore
        """
        if sub:
            tpl = self.template_with_sub
        else:
            tpl = self.template_without_sub
        return self.render_template(tpl, prefix_l=prefix_l,
                                    cat=cat, sub=sub or "", num=num)

    # ------------------------------------------------------------------
    # Validazione template
    # ------------------------------------------------------------------

    KNOWN_PATTERN = re.compile(r'\{L\}|\{CAT\}|\{SUB\}|\{NUM:\d+\}')
    BRACE_PATTERN = re.compile(r'\{[^}]+\}')

    @staticmethod
    def validate_template(template: str) -> list[str]:
        """
        Controlla il template e ritorna lista di variabili non riconosciute.
        Lista vuota = template valido.
        """
        unknown = []
        for m in re.findall(r'\{[^}]+\}', template):
            if not re.fullmatch(r'\{L\}|\{CAT\}|\{SUB\}|\{NUM:\d+\}', m):
                unknown.append(m)
        return unknown

    # ------------------------------------------------------------------
    # Helper per auto-codice categoria/sottocategoria (4 cifre numeriche)
    # ------------------------------------------------------------------

    def next_cat_code(self, existing_codes: list[str]) -> str:
        """
        Dato l'elenco dei codici categoria esistenti (es. ['0001','0002']),
        ritorna il prossimo codice 4-cifre disponibile (es. '0003').
        """
        return self._next_numeric_code(existing_codes, self.cat_code_length)

    def next_sub_code(self, existing_codes: list[str]) -> str:
        """
        Dato l'elenco dei codici sottocategoria esistenti,
        ritorna il prossimo codice 4-cifre disponibile.
        """
        return self._next_numeric_code(existing_codes, self.sub_code_length)

    @staticmethod
    def _next_numeric_code(existing_codes: list[str], length: int) -> str:
        """Trova il primo intero >= 1 non già presente tra i codici esistenti."""
        used = set()
        for c in existing_codes:
            try:
                used.add(int(c))
            except (ValueError, TypeError):
                pass
        n = 1
        while n in used:
            n += 1
        return str(n).zfill(length)

# =============================================================================
#  core/reg_manager.py  –  Gestione file .reg/.sldreg SolidWorks
#  Ispirato al parsing di LAUNCHER-SW, adattato per PDM-SW
# =============================================================================
from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class RestoreOptions:
    """Opzioni di ripristino selettivo per file .reg SolidWorks."""
    system_options: bool = True
    toolbar_layout: bool = True
    toolbar_mode: str = "all"           # "all" | "macro_only"
    keyboard_shortcuts: bool = True
    mouse_gestures: bool = True
    menu_customizations: bool = True
    saved_views: bool = True
    cleanup_before_import: bool = True

    def has_any_selection(self) -> bool:
        return any((
            self.system_options,
            self.toolbar_layout,
            self.keyboard_shortcuts,
            self.mouse_gestures,
            self.menu_customizations,
            self.saved_views,
        ))

    def describe(self) -> str:
        labels = []
        if self.system_options:
            labels.append("opzioni sistema")
        if self.toolbar_layout:
            labels.append("toolbar macro" if self.toolbar_mode == "macro_only"
                          else "layout toolbar")
        if self.keyboard_shortcuts:
            labels.append("tasti rapidi")
        if self.mouse_gestures:
            labels.append("gesti mouse")
        if self.menu_customizations:
            labels.append("menu")
        if self.saved_views:
            labels.append("viste salvate")
        return ", ".join(labels) if labels else "nessuna opzione"


# =========================================================================
#  Parsing file .reg / .sldreg
# =========================================================================

def _read_reg_text(file_path: Path) -> tuple[str, str]:
    """Legge un file .reg gestendo diversi encoding (UTF-16, UTF-8 BOM, ANSI)."""
    payload = file_path.read_bytes()
    if payload.startswith(b"\xff\xfe") or payload.startswith(b"\xfe\xff"):
        return payload.decode("utf-16"), "utf-16"
    if payload.startswith(b"\xef\xbb\xbf"):
        return payload.decode("utf-8-sig"), "utf-8-sig"
    for encoding in ("utf-8", "cp1252", "latin-1"):
        try:
            return payload.decode(encoding), encoding
        except Exception:
            continue
    return payload.decode("latin-1", errors="replace"), "latin-1"


def _section_from_registry_key(registry_key: str) -> str:
    """Estrae il nome sezione dalla chiave di registro SW."""
    parts = [p for p in registry_key.split("\\") if p]
    for idx, part in enumerate(parts):
        if part.upper().startswith("SOLIDWORKS "):
            if idx + 1 < len(parts):
                return parts[idx + 1]
            return "ROOT"
    return parts[-1] if parts else "ROOT"


def parse_reg_blocks(file_path: Path) -> tuple[
    list[str], list[tuple[str, str, list[str]]], str
]:
    """
    Parsa un file .reg/.sldreg in blocchi.
    Ritorna: (header_lines, [(section, registry_key, lines), ...], encoding)
    """
    text, encoding = _read_reg_text(file_path)
    raw_lines = text.splitlines()
    header_lines: list[str] = []
    blocks: list[tuple[str, str, list[str]]] = []
    current_lines: list[str] = []
    current_section = ""
    current_key = ""
    in_block = False

    for line in raw_lines:
        stripped = line.strip()
        if stripped.startswith("[") and stripped.endswith("]"):
            if in_block and current_lines:
                blocks.append((current_section, current_key, current_lines))
            key_name = stripped[1:-1].strip()
            current_section = _section_from_registry_key(key_name)
            current_key = key_name
            current_lines = [line]
            in_block = True
            continue
        if in_block:
            current_lines.append(line)
        else:
            header_lines.append(line)

    if in_block and current_lines:
        blocks.append((current_section, current_key, current_lines))
    if not header_lines:
        header_lines = ["Windows Registry Editor Version 5.00", ""]

    return header_lines, blocks, encoding


def list_reg_sections(file_path: Path) -> list[str]:
    """Elenca le sezioni uniche trovate nel file .reg."""
    _, blocks, _ = parse_reg_blocks(file_path)
    sections, seen = [], set()
    for section, _, _ in blocks:
        low = section.casefold()
        if low not in seen:
            seen.add(low)
            sections.append(section)
    return sections


# =========================================================================
#  Categorizzazione chiavi di registro
# =========================================================================

def _is_recent_key(key_name: str) -> bool:
    low = key_name.casefold()
    return "recent" in low or "pinned file list" in low


def registry_key_category(key_name: str) -> str:
    """Classifica una chiave di registro SW per categoria."""
    low = key_name.casefold()
    if _is_recent_key(key_name):
        return "recent"
    if "\\menu customizations" in low:
        return "menu_customizations"
    if "\\user interface\\saved views" in low:
        return "saved_views"
    if "\\user interface\\settings\\mouse gestures" in low:
        return "mouse_gestures"
    if "\\custom accelerators" in low:
        return "keyboard_shortcuts"
    if "\\user defined macros" in low:
        return "toolbar_macro"
    for token in ("\\user interface\\commandmanager",
                  "\\user interface\\api toolbars",
                  "\\user interface\\toolbars",
                  "\\user interface\\viewtools",
                  "\\simplified interface\\user interface\\viewtools",
                  "\\toolbars"):
        if token in low:
            return "toolbar_layout"
    return "system_options"


def _should_restore(category: str, options: RestoreOptions) -> bool:
    """Verifica se una categoria va ripristinata in base alle opzioni."""
    if category == "recent":
        return False
    mapping = {
        "system_options":      options.system_options,
        "toolbar_macro":       options.toolbar_layout,
        "toolbar_layout":      options.toolbar_layout and options.toolbar_mode == "all",
        "keyboard_shortcuts":  options.keyboard_shortcuts,
        "mouse_gestures":      options.mouse_gestures,
        "menu_customizations": options.menu_customizations,
        "saved_views":         options.saved_views,
    }
    return mapping.get(category, False)


# =========================================================================
#  Minimizzazione chiavi per cleanup
# =========================================================================

def _is_registry_parent(parent_key: str, child_key: str) -> bool:
    p = parent_key.rstrip("\\").casefold()
    c = child_key.rstrip("\\").casefold()
    return c == p or c.startswith(f"{p}\\")


def _minimize_cleanup_keys(raw_keys: list[str]) -> list[str]:
    unique: dict[str, str] = {}
    for key in raw_keys:
        n = key.strip()
        if n and n.casefold() not in unique:
            unique[n.casefold()] = n
    minimized = []
    for key in sorted(unique.values(),
                      key=lambda k: (k.count("\\"), len(k), k.casefold())):
        if not any(_is_registry_parent(e, key) for e in minimized):
            minimized.append(key)
    return minimized


def _collect_unsafe_parents(excluded_keys: list[str]) -> set[str]:
    unsafe: set[str] = set()
    for key in excluded_keys:
        parts = [p for p in key.strip().rstrip("\\").split("\\") if p]
        cursor = ""
        for part in parts:
            cursor = f"{cursor}\\{part}" if cursor else part
            unsafe.add(cursor.casefold())
    return unsafe


# =========================================================================
#  Import selettivo
# =========================================================================

def write_filtered_reg(
    source: Path,
    target: Path,
    options: RestoreOptions,
) -> tuple[bool, str]:
    """
    Scrive un file .reg filtrato con solo le categorie selezionate.
    Se cleanup_before_import, aggiunge righe di cancellazione [-key].
    Ritorna (ok, messaggio_errore).
    """
    if not options.has_any_selection():
        return False, "Seleziona almeno una voce da configurare."

    header, blocks, encoding = parse_reg_blocks(source)
    out = list(header)
    if out and out[-1] != "":
        out.append("")

    selected, excluded_keys, selected_keys = [], [], []
    for section, key, lines in blocks:
        cat = registry_key_category(key)
        if _should_restore(cat, options):
            selected.append((section, key, lines, cat))
            selected_keys.append(key)
        else:
            excluded_keys.append(key)

    if not selected:
        return False, "Nessuna chiave corrisponde alle opzioni selezionate."

    if options.cleanup_before_import:
        unsafe = _collect_unsafe_parents(excluded_keys)
        candidates = []
        seen = set()
        for key in selected_keys:
            n = key.strip().rstrip("\\")
            low = n.casefold()
            if low not in seen and low not in unsafe:
                seen.add(low)
                candidates.append(n)
        for key in _minimize_cleanup_keys(candidates):
            out.append(f"[-{key}]")
        if candidates:
            out.append("")

    for _, _, lines, _ in selected:
        out.extend(lines)
        if lines and lines[-1] != "":
            out.append("")

    text = "\r\n".join(out).rstrip() + "\r\n"
    target.write_text(text, encoding=encoding)
    return True, ""


def import_reg(
    file_path: Path,
    options: RestoreOptions | None = None,
) -> tuple[bool, str]:
    """
    Importa un file .reg/.sldreg nel registro di Windows con filtraggio
    selettivo per categoria. Ritorna (ok, messaggio).
    """
    opts = options or RestoreOptions()
    if not opts.has_any_selection():
        return False, "Seleziona almeno una voce da configurare."

    temp_path = None
    try:
        fd, temp_name = tempfile.mkstemp(prefix="pdm_reg_", suffix=".reg")
        os.close(fd)
        temp_path = Path(temp_name)

        ok, msg = write_filtered_reg(file_path, temp_path, opts)
        if not ok:
            return False, msg

        result = subprocess.run(
            ["reg", "import", str(temp_path)],
            capture_output=True, text=True, timeout=30,
        )
        if result.returncode == 0:
            mode = "pulizia + import" if opts.cleanup_before_import else "solo import"
            return True, f"Configurazione importata ({opts.describe()}, {mode})."
        err = (result.stderr or result.stdout or "").strip()
        return False, err or "Comando reg import terminato con errore."
    finally:
        if temp_path:
            try:
                temp_path.unlink(missing_ok=True)
            except Exception:
                pass


# =========================================================================
#  Rilevamento eseguibili
# =========================================================================

def detect_solidworks_exe() -> Path | None:
    """Cerca SLDWORKS.exe nel sistema."""
    from_path = shutil.which("SLDWORKS.exe")
    if from_path:
        p = Path(from_path)
        if p.is_file():
            return p
    for env in ("ProgramW6432", "ProgramFiles", "ProgramFiles(x86)"):
        val = os.environ.get(env, "").strip()
        if not val:
            continue
        sw_base = Path(val) / "SOLIDWORKS Corp"
        if not sw_base.is_dir():
            continue
        for pattern in ("SOLIDWORKS*/SLDWORKS.exe", "*/SLDWORKS.exe"):
            for c in sorted(sw_base.glob(pattern), reverse=True):
                if c.is_file():
                    return c
    return None


def detect_edrawings_exe() -> Path | None:
    """Cerca EModelView.exe (eDrawings) nel sistema."""
    for env in ("ProgramW6432", "ProgramFiles", "ProgramFiles(x86)"):
        val = os.environ.get(env, "").strip()
        if not val:
            continue
        base = Path(val)
        # Percorsi tipici eDrawings
        candidates = [
            base / "SOLIDWORKS Corp" / "eDrawings" / "EModelView.exe",
            base / "eDrawings" / "EModelView.exe",
        ]
        # Cerca anche con glob per versioni diverse
        sw_base = base / "SOLIDWORKS Corp"
        if sw_base.is_dir():
            for match in sw_base.glob("eDrawings*/EModelView.exe"):
                candidates.append(match)
        for c in candidates:
            if c.is_file():
                return c
    return None

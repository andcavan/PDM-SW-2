#!/usr/bin/env python
# =============================================================================
#  make_dist.py  –  Crea distribuzione portabile di PDM-SW
#
#  Uso:
#    python make_dist.py           # copia sorgenti in dist/PDM-SW-2/
#    python make_dist.py --zip     # copia + crea ZIP
#    python make_dist.py --clean   # elimina dist/ precedente prima di procedere
# =============================================================================
import shutil
import zipfile
import argparse
from pathlib import Path


# ---------------------------------------------------------------------------
#  Legge versione da config.py senza importarlo (evita dipendenza da PyQt6)
# ---------------------------------------------------------------------------
def _get_version() -> str:
    cfg = Path(__file__).parent / "config.py"
    for line in cfg.read_text(encoding="utf-8").splitlines():
        if line.strip().startswith("APP_VERSION"):
            return line.split("=")[1].strip().strip("\"'")
    return "0.0.0"


SRC_DIR = Path(__file__).parent.resolve()
VERSION = _get_version()
DIST_DIR = SRC_DIR / "dist" / "PDM-SW-2"

# ---------------------------------------------------------------------------
#  Regole di esclusione
# ---------------------------------------------------------------------------
EXCLUDE_DIRS = {
    ".venv", "__pycache__", "dist", ".git", ".idea", ".vscode",
}
EXCLUDE_FILES = {
    "local_config.json",    # dati personali del PC
    ".pdm_datadir",         # puntatore dati locali
    "sw_bridge_result.txt", # file runtime
    "sw_bridge.log",        # log runtime
}
EXCLUDE_EXTENSIONS = {".pyc", ".pyo", ".log"}


def should_exclude(rel: Path) -> bool:
    if any(part in EXCLUDE_DIRS for part in rel.parts):
        return True
    if rel.name in EXCLUDE_FILES:
        return True
    if rel.suffix in EXCLUDE_EXTENSIONS:
        return True
    return False


# ---------------------------------------------------------------------------
#  Copia sorgenti
# ---------------------------------------------------------------------------
def copy_sources(dest: Path) -> int:
    dest.mkdir(parents=True, exist_ok=True)
    count = 0
    for src_file in SRC_DIR.rglob("*"):
        if not src_file.is_file():
            continue
        rel = src_file.relative_to(SRC_DIR)
        if should_exclude(rel):
            continue
        dst = dest / rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src_file, dst)
        count += 1
    return count


# ---------------------------------------------------------------------------
#  Genera install.bat
# ---------------------------------------------------------------------------
def create_install_bat(dest: Path):
    content = (
        "@echo off\n"
        "REM =========================================================\n"
        "REM  install.bat  -  Installazione PDM-SW su nuovo PC\n"
        "REM  Richiede Python 3.10+ disponibile nel PATH di sistema\n"
        "REM =========================================================\n"
        "setlocal\n"
        "cd /d \"%~dp0\"\n"
        "\n"
        "echo.\n"
        "echo =========================================\n"
        f"echo   PDM-SW v{VERSION}  -  Installazione\n"
        "echo =========================================\n"
        "echo.\n"
        "\n"
        "REM -- Verifica Python --\n"
        "python --version >nul 2>&1\n"
        "if errorlevel 1 (\n"
        "    echo [ERRORE] Python non trovato nel PATH.\n"
        "    echo Scaricare Python 3.10+ da https://www.python.org\n"
        "    pause\n"
        "    exit /b 1\n"
        ")\n"
        "for /f \"tokens=*\" %%v in ('python --version 2^>^&1') do echo [OK] %%v\n"
        "\n"
        "REM -- Crea virtual environment --\n"
        "if not exist \".venv\\Scripts\\python.exe\" (\n"
        "    echo Creazione virtual environment...\n"
        "    python -m venv .venv\n"
        "    if errorlevel 1 (\n"
        "        echo [ERRORE] Impossibile creare il virtual environment.\n"
        "        pause\n"
        "        exit /b 1\n"
        "    )\n"
        "    echo [OK] Virtual environment creato.\n"
        ") else (\n"
        "    echo [OK] Virtual environment gia' presente.\n"
        ")\n"
        "\n"
        "REM -- Installa dipendenze --\n"
        "echo Installazione dipendenze (potrebbe richiedere qualche minuto)...\n"
        ".venv\\Scripts\\pip install -r requirements.txt\n"
        "if errorlevel 1 (\n"
        "    echo [ERRORE] Installazione dipendenze fallita.\n"
        "    pause\n"
        "    exit /b 1\n"
        ")\n"
        "echo [OK] Dipendenze installate.\n"
        "\n"
        "echo.\n"
        "echo =========================================\n"
        "echo   Installazione completata!\n"
        "echo   Avvia l'applicazione con start.bat\n"
        "echo =========================================\n"
        "echo.\n"
        "pause\n"
    )
    (dest / "install.bat").write_text(content, encoding="utf-8")


# ---------------------------------------------------------------------------
#  Genera start.bat
# ---------------------------------------------------------------------------
def create_start_bat(dest: Path):
    content = (
        "@echo off\n"
        "cd /d \"%~dp0\"\n"
        "if not exist \".venv\\Scripts\\python.exe\" (\n"
        "    echo Virtual environment non trovato.\n"
        "    echo Eseguire install.bat prima di avviare l'applicazione.\n"
        "    pause\n"
        "    exit /b 1\n"
        ")\n"
        "start \"\" \".venv\\Scripts\\pythonw.exe\" main.py\n"
    )
    (dest / "start.bat").write_text(content, encoding="utf-8")


# ---------------------------------------------------------------------------
#  Crea ZIP
# ---------------------------------------------------------------------------
def create_zip(dest_dir: Path) -> Path:
    zip_path = SRC_DIR / "dist" / f"PDM-SW-2_v{VERSION}.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for f in dest_dir.rglob("*"):
            if f.is_file():
                zf.write(f, f.relative_to(dest_dir.parent))
    return zip_path


# ---------------------------------------------------------------------------
#  Main
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Crea distribuzione portabile di PDM-SW")
    parser.add_argument("--zip", action="store_true", help="Crea anche il file ZIP")
    parser.add_argument("--clean", action="store_true", help="Elimina dist/ prima di procedere")
    args = parser.parse_args()

    if args.clean and DIST_DIR.exists():
        shutil.rmtree(DIST_DIR)
        print(f"[clean] Rimossa cartella: {DIST_DIR}")

    print(f"\nPDM-SW v{VERSION}")
    print(f"Destinazione: {DIST_DIR}\n")

    print("Copia sorgenti...", end=" ", flush=True)
    n = copy_sources(DIST_DIR)
    print(f"{n} file copiati.")

    print("Generazione install.bat e start.bat...", end=" ", flush=True)
    create_install_bat(DIST_DIR)
    create_start_bat(DIST_DIR)
    print("OK")

    if args.zip:
        print("Creazione ZIP...", end=" ", flush=True)
        zip_path = create_zip(DIST_DIR)
        size_mb = zip_path.stat().st_size / 1_048_576
        print(f"OK  ({size_mb:.1f} MB)  →  {zip_path}")

    print(f"\nDistribuzione pronta in:\n  {DIST_DIR}\n")


if __name__ == "__main__":
    main()

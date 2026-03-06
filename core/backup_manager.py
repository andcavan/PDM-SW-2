# =============================================================================
#  core/backup_manager.py  –  Backup automatico del database PDM
# =============================================================================
import shutil
from datetime import datetime
from pathlib import Path


class BackupManager:
    """
    Gestisce i backup del database pdm.db nella stessa cartella condivisa.
    I backup vengono salvati in <shared_root>/database/backups/
    con nome pdm_YYYYMMDD_HHMMSS.db.
    Mantiene solo gli ultimi `keep` backup (default 10).
    """

    def __init__(self, db_path: Path, keep: int = 10):
        self.db_path    = db_path
        self.backup_dir = db_path.parent / "backups"
        self.keep       = keep

    # ------------------------------------------------------------------
    def create(self) -> Path:
        """
        Crea un backup immediato del database.
        Ritorna il percorso del file di backup creato.
        Lancia eccezione se il database sorgente non esiste.
        """
        if not self.db_path.exists():
            raise FileNotFoundError(f"Database non trovato: {self.db_path}")

        self.backup_dir.mkdir(parents=True, exist_ok=True)

        ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
        dest = self.backup_dir / f"pdm_{ts}.db"
        shutil.copy2(self.db_path, dest)
        self._purge_old()
        return dest

    # ------------------------------------------------------------------
    def list_backups(self) -> list[Path]:
        """Restituisce la lista dei backup ordinata dal più recente."""
        if not self.backup_dir.exists():
            return []
        files = sorted(
            self.backup_dir.glob("pdm_*.db"),
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
        return files

    # ------------------------------------------------------------------
    def restore(self, backup_path: Path):
        """
        Ripristina il database da un backup.
        Il database attuale viene rinominato con suffisso .bak prima del ripristino.
        """
        if not backup_path.exists():
            raise FileNotFoundError(f"Backup non trovato: {backup_path}")

        # Salva copia di sicurezza del DB corrente
        if self.db_path.exists():
            bak = self.db_path.with_suffix(".bak")
            shutil.copy2(self.db_path, bak)

        shutil.copy2(backup_path, self.db_path)

    # ------------------------------------------------------------------
    def _purge_old(self):
        """Elimina i backup più vecchi mantenendo solo gli ultimi `keep`."""
        files = self.list_backups()
        for old in files[self.keep:]:
            try:
                old.unlink()
            except Exception:
                pass

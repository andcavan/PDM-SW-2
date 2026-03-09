# =============================================================================
#  ui/regen_thumbnails_dialog.py  –  Rigenerazione massiva thumbnail
# =============================================================================
import sys
import subprocess
import tempfile
import os
from pathlib import Path

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QProgressBar, QTextEdit, QCheckBox, QMessageBox
)
from PyQt6.QtCore import QThread, pyqtSignal

from ui.session import session

# Script subprocess per generare UNA thumbnail (isolato, killabile)
_WORKER_SCRIPT = Path(__file__).parent.parent / "core" / "thumb_worker.py"
_TIMEOUT_SEC   = 35   # secondi massimi per file prima di uccidere il processo


class _RegenWorker(QThread):
    """Thread che processa i documenti in sequenza, un subprocess per file."""
    progress    = pyqtSignal(int, int, str)  # (current, total, label)
    log         = pyqtSignal(str)
    finished_ok = pyqtSignal(int, int)       # (generati, saltati)

    def __init__(self, docs: list, sp, overwrite: bool):
        super().__init__()
        self._docs      = docs
        self._sp        = sp
        self._overwrite = overwrite
        self._abort     = False
        self._proc      = None   # subprocess corrente

    def abort(self):
        self._abort = True
        if self._proc:
            try:
                self._proc.kill()
            except Exception:
                pass

    def run(self):
        if not _WORKER_SCRIPT.exists():
            self.log.emit(f"[ERRORE] Script worker non trovato:\n{_WORKER_SCRIPT}")
            self.finished_ok.emit(0, len(self._docs))
            return

        python_exe = Path(sys.executable)
        ext_map    = {"Parte": ".SLDPRT", "Assieme": ".SLDASM", "Disegno": ".SLDDRW"}
        generated  = 0
        skipped    = 0
        total      = len(self._docs)

        for i, doc in enumerate(self._docs, 1):
            if self._abort:
                self.log.emit("Operazione interrotta dall'utente.")
                break

            code  = doc.get("code", "")
            rev   = doc.get("revision", "")
            label = f"{code} rev.{rev}"
            self.progress.emit(i, total, label)

            suffix = "_DRW" if doc.get("doc_type") == "Disegno" else ""
            dest = self._sp.thumbnails / f"{code}_{rev}{suffix}.png"

            if dest.exists() and not self._overwrite:
                skipped += 1
                continue

            # Trova il file sorgente nell'archivio
            src_file = None
            file_name = doc.get("file_name") or (
                code + ext_map.get(doc.get("doc_type", ""), "")
            )
            try:
                arch_path = self._sp.archive_path(code, rev)
                candidate = arch_path / file_name
                if candidate.exists():
                    src_file = candidate
            except Exception:
                pass

            if src_file is None and doc.get("archive_path"):
                candidate = self._sp.root / doc["archive_path"]
                if candidate.exists():
                    src_file = candidate

            if not src_file:
                self.log.emit(f"  [SALTA] {label} — file non trovato in archivio")
                skipped += 1
                continue

            # Lancia subprocess isolato con timeout.
            # Usiamo un file temporaneo per stdout/stderr invece di PIPE:
            # i surrogate COM (dllhost.exe) ereditano i PIPE-handle e tengono
            # communicate() bloccato anche dopo la fine del processo principale.
            tmp_fd, tmp_log = tempfile.mkstemp(suffix='.txt', prefix='thumb_')
            os.close(tmp_fd)
            try:
                with open(tmp_log, 'w') as log_f:
                    self._proc = subprocess.Popen(
                        [str(python_exe), str(_WORKER_SCRIPT),
                         str(src_file), str(dest)],
                        stdout=log_f,
                        stderr=log_f,
                        creationflags=subprocess.CREATE_NO_WINDOW,
                    )
                try:
                    self._proc.wait(timeout=_TIMEOUT_SEC)
                    out = Path(tmp_log).read_text(errors='replace').strip()

                    if self._proc.returncode == 0 and dest.exists():
                        generated += 1
                        self.log.emit(f"  [OK] {label}")
                    else:
                        detail = out or f"exit code {self._proc.returncode}"
                        self.log.emit(f"  [FALLITO] {label}: {detail}")
                        skipped += 1

                except subprocess.TimeoutExpired:
                    self._proc.kill()
                    self._proc.wait()
                    self.log.emit(
                        f"  [TIMEOUT] {label} — annullato dopo {_TIMEOUT_SEC} s"
                    )
                    skipped += 1

            except Exception as e:
                self.log.emit(f"  [ERRORE] {label}: {e}")
                skipped += 1
            finally:
                self._proc = None
                try:
                    os.unlink(tmp_log)
                except OSError:
                    pass

        self.finished_ok.emit(generated, skipped)


class RegenThumbnailsDialog(QDialog):
    """
    Dialog per la rigenerazione massiva delle thumbnail.
    Accessibile da Strumenti → Rigenera anteprime.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Rigenera anteprime documenti")
        self.setMinimumSize(540, 420)
        self._worker: _RegenWorker | None = None
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        info = QLabel(
            "Estrae l'anteprima embedded dal file SolidWorks (OLE) oppure, se non\n"
            "disponibile, la genera tramite eDrawings. Ogni file ha un timeout di "
            f"{_TIMEOUT_SEC} s."
        )
        info.setWordWrap(True)
        layout.addWidget(info)

        self.chk_overwrite = QCheckBox("Sovrascrivi anteprime già esistenti")
        self.chk_overwrite.setChecked(False)
        layout.addWidget(self.chk_overwrite)

        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        layout.addWidget(self.progress_bar)

        self.lbl_status = QLabel("Pronto.")
        self.lbl_status.setStyleSheet("color: #888; font-size: 11px;")
        layout.addWidget(self.lbl_status)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setPlaceholderText("Il log apparirà qui…")
        layout.addWidget(self.log_box)

        btn_row = QHBoxLayout()
        self.btn_start = QPushButton("▶  Avvia")
        self.btn_start.setDefault(True)
        self.btn_start.clicked.connect(self._start)

        self.btn_abort = QPushButton("■  Interrompi")
        self.btn_abort.setEnabled(False)
        self.btn_abort.clicked.connect(self._abort)

        self.btn_close = QPushButton("Chiudi")
        self.btn_close.clicked.connect(self.close)

        btn_row.addWidget(self.btn_start)
        btn_row.addWidget(self.btn_abort)
        btn_row.addStretch()
        btn_row.addWidget(self.btn_close)
        layout.addLayout(btn_row)

    def _log(self, msg: str):
        self.log_box.append(msg)
        sb = self.log_box.verticalScrollBar()
        if sb:
            sb.setValue(sb.maximum())

    def _start(self):
        if not session.sp or not session.files:
            QMessageBox.warning(
                self, "Percorso rete non configurato",
                "Configurare il percorso rete in Strumenti → Configurazione percorso rete."
            )
            return

        docs = [d for d in session.files.search_documents() if d.get("archive_path")]
        if not docs:
            QMessageBox.information(
                self, "Nessun documento",
                "Nessun documento con file archiviato trovato nel database."
            )
            return

        overwrite = self.chk_overwrite.isChecked()
        self.progress_bar.setMaximum(len(docs))
        self.progress_bar.setValue(0)
        self.log_box.clear()
        self._log(
            f"Avvio: {len(docs)} documenti. "
            f"Sovrascrittura: {'sì' if overwrite else 'no'}. "
            f"Timeout per file: {_TIMEOUT_SEC} s."
        )
        self.lbl_status.setText(f"Elaborazione di {len(docs)} documenti…")

        self.btn_start.setEnabled(False)
        self.btn_abort.setEnabled(True)
        self.chk_overwrite.setEnabled(False)

        self._worker = _RegenWorker(docs, session.sp, overwrite)
        self._worker.progress.connect(self._on_progress)
        self._worker.log.connect(self._log)
        self._worker.finished_ok.connect(self._on_finished)
        self._worker.start()

    def _abort(self):
        if self._worker:
            self._worker.abort()
            self.btn_abort.setEnabled(False)
            self.lbl_status.setText("Interruzione in corso…")

    def _on_progress(self, current: int, total: int, label: str):
        self.progress_bar.setValue(current)
        self.lbl_status.setText(f"[{current}/{total}]  {label}")

    def _on_finished(self, generated: int, skipped: int):
        self.progress_bar.setValue(self.progress_bar.maximum())
        msg = f"Completato: {generated} anteprime generate, {skipped} saltate/fallite."
        self.lbl_status.setText(msg)
        self._log(f"\n{msg}")
        self.btn_start.setEnabled(True)
        self.btn_abort.setEnabled(False)
        self.chk_overwrite.setEnabled(True)
        self._worker = None

    def closeEvent(self, event):
        if self._worker and self._worker.isRunning():
            self._worker.abort()
            self._worker.wait(4000)
        super().closeEvent(event)

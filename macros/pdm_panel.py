#!/usr/bin/env python
# =============================================================================
#  macros/pdm_panel.py  –  Pannello Qt rapido per SolidWorks  v1.0
#
#  Lanciato dalla macro VBA:
#    pythonw pdm_panel.py [percorso_file_attivo]
#
#  Mostra: stato documento, pulsanti azione, form genera-codice
# =============================================================================
from __future__ import annotations
import sys
import logging
import socket
from pathlib import Path
from typing import Optional, TYPE_CHECKING

if TYPE_CHECKING:
    from core.checkout_manager import CheckoutManager
    from core.file_manager import FileManager
    from core.properties_manager import PropertiesManager
    from core.coding_manager import CodingManager
    from core.asm_manager import AsmManager
    from core.database import Database
    from config import SharedPaths

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

from PyQt6.QtWidgets import (
    QApplication, QDialog, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QPushButton, QFrame, QMessageBox, QLineEdit,
    QComboBox, QGroupBox, QSizePolicy,
)
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QFont, QIcon

# ---------------------------------------------------------------------------
LOG_FILE = ROOT / "macros" / "sw_bridge.log"

def _setup_log():
    logging.basicConfig(
        filename=str(LOG_FILE),
        level=logging.DEBUG,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    # Forza flush immediato ad ogni scrittura
    for h in logging.getLogger().handlers:
        if hasattr(h, 'stream'):
            h.stream.reconfigure(line_buffering=True) if hasattr(h.stream, 'reconfigure') else None

_setup_log()

# ---------------------------------------------------------------------------
#  Costanti colore (in linea con DARK_THEME)
# ---------------------------------------------------------------------------
STATE_COLORS = {
    "In Lavorazione": "#89b4fa",
    "Rilasciato":     "#a6e3a1",
    "In Revisione":   "#fab387",
    "Obsoleto":       "#585b70",
}
TYPE_ICONS = {
    "Parte":   "⚙",
    "Assieme": "🔩",
    "Disegno": "📐",
}
EXT_TO_TYPE = {
    ".SLDPRT": "Parte",
    ".SLDASM": "Assieme",
    ".SLDDRW": "Disegno",
}

STYLE = """
QDialog, QWidget {
    background-color: #1e1e2e;
    color: #cdd6f4;
    font-family: 'Segoe UI', Arial, sans-serif;
    font-size: 12px;
}
QGroupBox {
    border: 1px solid #45475a;
    border-radius: 5px;
    margin-top: 6px;
    padding-top: 4px;
    color: #89b4fa;
    font-weight: bold;
}
QGroupBox::title { subcontrol-origin: margin; left: 8px; padding: 0 4px; }
QPushButton {
    background-color: #313244;
    color: #cdd6f4;
    border: 1px solid #45475a;
    border-radius: 5px;
    padding: 6px 10px;
    min-width: 130px;
    text-align: left;
    font-size: 12px;
}
QPushButton:hover { background-color: #45475a; border-color: #89b4fa; }
QPushButton:pressed { background-color: #585b70; }
QPushButton:disabled { color: #585b70; border-color: #313244; }
QPushButton#btn_checkout  { border-left: 3px solid #89b4fa; }
QPushButton#btn_checkin   { border-left: 3px solid #a6e3a1; }
QPushButton#btn_undo      { border-left: 3px solid #fab387; }
QPushButton#btn_props_in  { border-left: 3px solid #cba6f7; }
QPushButton#btn_props_out { border-left: 3px solid #cba6f7; }
QPushButton#btn_open      { border-left: 3px solid #f38ba8; }
QPushButton#btn_generate  {
    background-color: #2a2a3e;
    border: 1px dashed #a6e3a1;
    color: #a6e3a1;
    font-weight: bold;
    min-width: 270px;
    text-align: center;
}
QPushButton#btn_generate:hover { background-color: #313244; border-color: #cdd6f4; }
QPushButton#btn_refresh {
    background-color: transparent;
    border: none;
    color: #585b70;
    min-width: 26px;
    padding: 2px;
    font-size: 14px;
}
QPushButton#btn_refresh:hover { color: #89b4fa; }
QLineEdit, QComboBox {
    background-color: #313244;
    border: 1px solid #45475a;
    border-radius: 4px;
    padding: 4px 8px;
    color: #cdd6f4;
}
QLineEdit:focus, QComboBox:focus { border-color: #89b4fa; }
QComboBox::drop-down { border: none; width: 20px; }
QComboBox QAbstractItemView {
    background-color: #313244;
    border: 1px solid #45475a;
    selection-background-color: #45475a;
}
QFrame#separator {
    color: #45475a;
    max-height: 1px;
    background-color: #45475a;
}
"""


# ===========================================================================
#  Sessione leggera (senza Qt session globale)
# ===========================================================================
def _load_bare_session():
    """Apre DB e trova l'utente per hostname. Ritorna (db, sp, user, coding, files, checkout, props, asm)."""
    from config import load_local_config, SharedPaths
    from core.database import Database
    from core.checkout_manager import CheckoutManager
    from core.file_manager import FileManager
    from core.properties_manager import PropertiesManager
    from core.coding_manager import CodingManager
    from core.asm_manager import AsmManager

    cfg = load_local_config()
    shared_root = cfg.get("shared_root", "")
    if not shared_root:
        raise ValueError(
            "Percorso condiviso non configurato.\nAprire PDM-SW e configurare il percorso."
        )

    sp = SharedPaths(shared_root)
    db = Database(sp.db_file, sp.db_lock_file)

    hostname = socket.gethostname()
    user = (
        db.fetchone("SELECT * FROM users WHERE workstation=? AND active=1", (hostname,))
        or db.fetchone("SELECT * FROM users WHERE role='Amministratore' AND active=1")
        or db.fetchone("SELECT * FROM users WHERE active=1")
    )
    if not user:
        raise ValueError("Nessun utente trovato nel database PDM.")

    coding   = CodingManager(db)
    files    = FileManager(db, sp, user)
    checkout = CheckoutManager(db, sp, user)
    props    = PropertiesManager(db)
    asm      = AsmManager(db)

    return db, sp, user, coding, files, checkout, props, asm


def _find_doc(db, file_path: str, user=None, need_lock=False):
    """Trova documento PDM dal percorso file."""
    from config import SW_EXTENSIONS
    fp = Path(file_path)
    code = fp.stem
    ext  = fp.suffix.upper()
    doc_type = SW_EXTENSIONS.get(ext, "")

    if need_lock and user:
        doc = db.fetchone(
            "SELECT * FROM documents WHERE code=? AND doc_type=? AND is_locked=1 AND locked_by=? "
            "ORDER BY revision DESC",
            (code, doc_type, user["id"]),
        )
        if doc:
            return doc

    doc = db.fetchone(
        "SELECT * FROM documents WHERE code=? AND doc_type=? AND state != 'Obsoleto' "
        "ORDER BY revision DESC",
        (code, doc_type),
    )
    return doc or db.fetchone("SELECT * FROM documents WHERE file_name=?", (fp.name,))


# ===========================================================================
#  Dialogo genera codice
# ===========================================================================
class CreateCodeDialog(QDialog):
    """Sotto-dialogo per assegnare un codice PDM a un file SW."""

    def __init__(self, file_path: str, db, sp, user, coding, files, parent=None):
        super().__init__(parent)
        self.file_path = file_path
        self.db        = db
        self.sp        = sp
        self.user      = user
        self.coding    = coding
        self.files     = files
        self.created_doc_id: Optional[int] = None

        fp         = Path(file_path)
        self.code  = fp.stem
        self.ext   = fp.suffix.upper()
        self.doc_type = EXT_TO_TYPE.get(self.ext, "Parte")

        self.setWindowTitle("Genera Codice PDM")
        self.setMinimumWidth(360)
        self.setStyleSheet(STYLE)
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(14, 14, 14, 14)

        # --- Info file ---
        info = QLabel(
            f"File: <b>{Path(self.file_path).name}</b><br>"
            f"Tipo documento: <b>{self.doc_type}</b>"
        )
        info.setTextFormat(Qt.TextFormat.RichText)
        info.setWordWrap(True)
        layout.addWidget(info)

        # --- Form ---
        form_group = QGroupBox("Dati documento")
        form = QGridLayout(form_group)
        form.setHorizontalSpacing(10)
        form.setVerticalSpacing(8)

        # Macchina
        form.addWidget(QLabel("Macchina:"), 0, 0)
        self.cmb_machine = QComboBox()
        self.cmb_machine.setMinimumWidth(200)
        form.addWidget(self.cmb_machine, 0, 1)

        # Gruppo
        form.addWidget(QLabel("Gruppo:"), 1, 0)
        self.cmb_group = QComboBox()
        form.addWidget(self.cmb_group, 1, 1)

        # Livello (per ASM: Macchina/Gruppo/Sottogruppo)
        form.addWidget(QLabel("Livello:"), 2, 0)
        self.cmb_level = QComboBox()
        form.addWidget(self.cmb_level, 2, 1)

        # Titolo
        form.addWidget(QLabel("Titolo:"), 3, 0)
        self.txt_title = QLineEdit()
        self.txt_title.setPlaceholderText("Descrizione breve…")
        form.addWidget(self.txt_title, 3, 1)

        layout.addWidget(form_group)

        # Preview codice
        self.lbl_preview = QLabel("Codice: —")
        self.lbl_preview.setStyleSheet(
            "font-size:14px; font-weight:bold; color:#a6e3a1; padding:4px;"
        )
        self.lbl_preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.lbl_preview)

        # Bottoni
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_cancel = QPushButton("Annulla")
        btn_cancel.clicked.connect(self.reject)
        self.btn_create = QPushButton("✅  Crea Documento")
        self.btn_create.setObjectName("btn_checkin")
        self.btn_create.clicked.connect(self._create)
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(self.btn_create)
        layout.addLayout(btn_row)

        # Carica dati e collega segnali
        self._populate_machines()
        self.cmb_machine.currentIndexChanged.connect(self._on_machine_changed)
        self.cmb_group.currentIndexChanged.connect(self._update_preview)
        self.cmb_level.currentIndexChanged.connect(self._update_preview)
        self._on_machine_changed()

    # -- helpers -----------------------------------------------------------
    def _populate_machines(self):
        self.cmb_machine.clear()
        machines = self.coding.get_machines()
        for m in machines:
            self.cmb_machine.addItem(f"{m['code']}  –  {m['description']}", m["id"])
        if not machines:
            self.cmb_machine.addItem("(nessuna macchina)", None)

    def _on_machine_changed(self, _=None):
        self.cmb_group.clear()
        mid = self.cmb_machine.currentData()
        if mid is None:
            return
        groups = self.coding.get_groups(mid)
        for g in groups:
            self.cmb_group.addItem(f"{g['code']}  –  {g['description']}", g["id"])
        if not groups:
            self.cmb_group.addItem("(nessun gruppo)", None)
        self._update_level_options()
        self._update_preview()

    def _update_level_options(self):
        self.cmb_level.blockSignals(True)
        self.cmb_level.clear()
        if self.doc_type in ("Parte",):
            self.cmb_level.addItem("Livello 2 – Parte",     ("PRT", 2))
        elif self.doc_type == "Disegno":
            self.cmb_level.addItem("Livello 2 – Disegno",   ("PRT", 2))
        else:  # Assieme
            self.cmb_level.addItem("Livello 2 – Sottogruppo",  ("ASM_SUB", 2))
            self.cmb_level.addItem("Livello 1 – Gruppo ASM",   ("ASM_GRP", 1))
            self.cmb_level.addItem("Livello 0 – Macchina ASM", ("ASM_MACH", 0))
        self.cmb_level.blockSignals(False)
        self._update_preview()

    def _update_preview(self, _=None):
        mid = self.cmb_machine.currentData()
        gid = self.cmb_group.currentData()
        lv_data = self.cmb_level.currentData()
        if not mid or not lv_data:
            self.lbl_preview.setText("Codice: —")
            return
        subtype, level = lv_data
        try:
            if subtype == "ASM_MACH":
                code = self.coding.preview_code(0, "ASM", mid, gid)
            elif subtype == "ASM_GRP":
                code = self.coding.preview_code(1, "ASM", mid, gid)
            elif subtype == "ASM_SUB":
                code = self.coding.preview_code(2, "ASM", mid, gid)
            else:
                code = self.coding.preview_code(2, "PRT", mid, gid)
            self.lbl_preview.setText(f"Codice: {code}  rev. 00")
        except Exception as e:
            self.lbl_preview.setText(f"—  ({e})")

    def _create(self):
        logging.info("_create() avviato — file_path=%r ext=%r doc_type=%r",
                     self.file_path, self.ext, self.doc_type)
        title = self.txt_title.text().strip()
        if not title:
            QMessageBox.warning(self, "Campo obbligatorio", "Inserire il titolo.")
            return

        mid = self.cmb_machine.currentData()
        gid = self.cmb_group.currentData()
        lv_data = self.cmb_level.currentData()
        if not mid or not lv_data:
            QMessageBox.warning(self, "Selezione", "Selezionare macchina e livello.")
            return

        subtype, level = lv_data
        try:
            if subtype == "ASM_MACH":
                code = self.coding.next_code_liv0(mid)
            elif subtype == "ASM_GRP":
                code = self.coding.next_code_liv1(mid, gid)
            elif subtype == "ASM_SUB":
                code = self.coding.next_code_liv2_part(mid, gid)
            else:
                code = self.coding.next_code_liv2_subgroup(mid, gid)

            from config import load_local_config
            import shutil, socket
            ws_root = Path(load_local_config().get("sw_workspace", "") or "")
            if not ws_root or not ws_root.is_dir():
                QMessageBox.critical(
                    self, "Workspace non configurata",
                    "Configurare la workspace SolidWorks nelle impostazioni."
                )
                return

            # --- Determina estensione e salva il file ---
            ext = self.ext  # può essere vuota se il file non era ancora salvato

            sw_doc = None
            try:
                import win32com.client
                sw = win32com.client.GetActiveObject("SldWorks.Application")
                sw_doc = sw.ActiveDoc
                if sw_doc and not ext:
                    # Ricava l'estensione dal tipo documento SW
                    sw_type = sw_doc.GetType()   # 1=Parte, 2=Assieme, 3=Disegno
                    ext = {1: ".SLDPRT", 2: ".SLDASM", 3: ".SLDDRW"}.get(sw_type, ".SLDPRT")
            except Exception as sw_err:
                logging.warning("SW COM non disponibile: %s", sw_err)

            if not ext:
                QMessageBox.critical(
                    self, "Tipo file sconosciuto",
                    "Impossibile determinare il tipo del file.\n\n"
                    "Salvare il file in SolidWorks almeno una volta prima di assegnare il codice PDM."
                )
                return

            # Tipo documento coerente con l'estensione effettiva
            doc_type = EXT_TO_TYPE.get(ext, self.doc_type)
            dest_ws = ws_root / f"{code}{ext}"

            # --- Save As via COM oppure copia fisica ---
            saved = False
            logging.info("Tentativo salvataggio: sw_doc=%s dest_ws=%s", bool(sw_doc), dest_ws)
            if sw_doc:
                try:
                    # Se destinazione == sorgente non serve SaveAs
                    src_path = Path(self.file_path) if self.file_path else None
                    if src_path and src_path.resolve() == dest_ws.resolve():
                        logging.info("File già in workspace con nome corretto, skip SaveAs")
                        saved = True
                    else:
                        sw_doc.SaveAs3(str(dest_ws), 0, 0)
                        if dest_ws.exists():
                            saved = True
                            logging.info("SW SaveAs3 OK → %s", dest_ws)
                        else:
                            logging.warning("SaveAs3 non ha creato il file")
                except Exception as e:
                    logging.warning("SaveAs3 fallita: %s", e)

            if not saved:
                src = Path(self.file_path) if self.file_path else None
                if not src or not src.exists():
                    QMessageBox.critical(
                        self, "File non trovabile",
                        "Nessun file da salvare.\n\n"
                        "Aprire e salvare il file in SolidWorks prima di assegnare il codice."
                    )
                    return
                shutil.copy2(src, dest_ws)
                logging.info("File copiato in workspace → %s", dest_ws)

            if not dest_ws.exists():
                QMessageBox.critical(
                    self, "Errore salvataggio",
                    f"Il file non è stato creato nella workspace:\n{dest_ws}"
                )
                return

            # --- Crea documento nel DB ---
            doc_id = self.files.create_document(
                code=code, revision="00", doc_type=doc_type,
                title=title, machine_id=mid, group_id=gid,
                doc_level=level if level > 0 else 2,
            )
            logging.info("Documento creato: code=%s id=%d ext=%s type=%s", code, doc_id, ext, doc_type)

            # --- Archivia direttamente (senza passare per import_file) ---
            try:
                arch_dir = self.sp.archive_path(code, "00")
                arch_dir.mkdir(parents=True, exist_ok=True)
                arch_file = arch_dir / f"{code}{ext}"
                shutil.copy2(str(dest_ws), str(arch_file))
                rel_path = str(arch_file.relative_to(self.sp.root))
                uid = self.user["id"]
                self.db.execute(
                    """UPDATE documents
                       SET file_name=?, file_ext=?, archive_path=?,
                           modified_by=?, modified_at=datetime('now')
                       WHERE id=?""",
                    (arch_file.name, ext, rel_path, uid, doc_id),
                )
                logging.info("File archiviato → %s", arch_file)
            except Exception as arch_err:
                # Rollback DB: elimina record appena creato
                try:
                    self.db.execute("DELETE FROM documents WHERE id=?", (doc_id,))
                except Exception:
                    pass
                raise RuntimeError(f"Archiviazione fallita: {arch_err}") from arch_err

            # --- Lock come checkout ---
            uid = self.user["id"]
            self.db.execute(
                """UPDATE documents
                   SET is_locked=1, locked_by=?, locked_at=datetime('now'),
                       locked_ws=?
                   WHERE id=?""",
                (uid, socket.gethostname(), doc_id),
            )
            self.db.execute(
                """INSERT INTO checkout_log
                   (document_id, user_id, action, workstation, workspace_path)
                   VALUES (?,?,'checkout',?,?)""",
                (doc_id, uid, socket.gethostname(), str(dest_ws)),
            )
            existing = self.db.fetchone(
                "SELECT id FROM workspace_files WHERE document_id=? AND user_id=?",
                (doc_id, uid),
            )
            if existing:
                self.db.execute(
                    """UPDATE workspace_files
                       SET role='checkout', workspace_path=?, copied_at=datetime('now')
                       WHERE id=?""",
                    (str(dest_ws), existing["id"]),
                )
            else:
                self.db.execute(
                    """INSERT INTO workspace_files
                       (document_id, user_id, role, workspace_path)
                       VALUES (?,?,'checkout',?)""",
                    (doc_id, uid, str(dest_ws)),
                )

            self.created_doc_id = doc_id
            self.created_path   = dest_ws
            QMessageBox.information(
                self, "Documento creato",
                f"Codice: {code}  rev.00  [{self.doc_type}]\n"
                f"Titolo: {title}\n\n"
                f"Workspace:  {dest_ws}\n"
                f"Archivio:   {arch_file}"
            )
            self.accept()

        except Exception as e:
            logging.error("Errore creazione documento: %s", e, exc_info=True)
            QMessageBox.critical(self, "Errore", str(e))


# ===========================================================================
#  Pannello principale
# ===========================================================================
class PDMPanel(QDialog):
    """Pannello flottante con azioni PDM per il file SW attivo."""

    def __init__(self, file_path: str):
        super().__init__()
        self.file_path = file_path
        self.doc:      Optional[dict]             = None
        self.db:       Optional[Database]          = None
        self.sp:       Optional[SharedPaths]       = None
        self.user:     Optional[dict]              = None
        self.coding:   Optional[CodingManager]     = None
        self.files:    Optional[FileManager]       = None
        self.checkout: Optional[CheckoutManager]   = None
        self.props:    Optional[PropertiesManager] = None
        self.asm:      Optional[AsmManager]        = None

        self.setWindowTitle("PDM-SW")
        self.setMinimumWidth(300)
        self.setMaximumWidth(380)
        self.setStyleSheet(STYLE)
        self.setWindowFlags(
            Qt.WindowType.WindowStaysOnTopHint |
            Qt.WindowType.Dialog
        )

        try:
            (self.db, self.sp, self.user,
             self.coding, self.files,
             self.checkout, self.props, self.asm) = _load_bare_session()
            self.session_error = None
        except Exception as e:
            self.session_error = str(e)
            logging.error("PDMPanel - sessione fallita: %s", e)

        self._build_ui()
        if not self.session_error:
            self._refresh()

    # ------------------------------------------------------------------
    def _build_ui(self):
        self._main_layout = QVBoxLayout(self)
        self._main_layout.setSpacing(8)
        self._main_layout.setContentsMargins(12, 10, 12, 12)

        if self.session_error:
            err_lbl = QLabel(f"⚠ Sessione non disponibile:\n{self.session_error}")
            err_lbl.setStyleSheet("color:#f38ba8;")
            err_lbl.setWordWrap(True)
            self._main_layout.addWidget(err_lbl)
            return

        # --- Header file info ---
        self._grp_info = QGroupBox("Documento")
        info_layout = QVBoxLayout(self._grp_info)
        info_layout.setSpacing(3)

        # Riga nome file + refresh
        hdr_row = QHBoxLayout()
        self._lbl_filename = QLabel("—")
        self._lbl_filename.setStyleSheet("font-weight:bold; font-size:12px;")
        self._lbl_filename.setWordWrap(True)
        hdr_row.addWidget(self._lbl_filename, 1)

        btn_refresh = QPushButton("⟳")
        btn_refresh.setObjectName("btn_refresh")
        btn_refresh.setToolTip("Aggiorna stato")
        btn_refresh.clicked.connect(self._refresh)
        hdr_row.addWidget(btn_refresh, 0)
        info_layout.addLayout(hdr_row)

        # Tipo + revisione
        self._lbl_meta = QLabel()
        self._lbl_meta.setStyleSheet("color:#a6adc8; font-size:11px;")
        info_layout.addWidget(self._lbl_meta)

        # Stato
        self._lbl_state = QLabel()
        self._lbl_state.setTextFormat(Qt.TextFormat.RichText)
        info_layout.addWidget(self._lbl_state)

        # Lock
        self._lbl_lock = QLabel()
        self._lbl_lock.setWordWrap(True)
        info_layout.addWidget(self._lbl_lock)

        self._main_layout.addWidget(self._grp_info)

        # --- Pulsanti azioni ---
        self._grp_actions = QGroupBox("Azioni")
        grid = QGridLayout(self._grp_actions)
        grid.setSpacing(6)

        self.btn_co = QPushButton("📤  Checkout")
        self.btn_co.setObjectName("btn_checkout")
        self.btn_co.setToolTip("Scarica in workspace e blocca per modifica")
        self.btn_co.clicked.connect(self._do_checkout)

        self.btn_ci = QPushButton("📥  Check-in")
        self.btn_ci.setObjectName("btn_checkin")
        self.btn_ci.setToolTip("Archivia file modificato e rilascia il lock")
        self.btn_ci.clicked.connect(self._do_checkin)

        self.btn_undo = QPushButton("↩  Annulla CO")
        self.btn_undo.setObjectName("btn_undo")
        self.btn_undo.setToolTip("Rilascia il lock senza archiviare le modifiche")
        self.btn_undo.clicked.connect(self._do_undo_checkout)

        self.btn_consult = QPushButton("👁  Consultazione")
        self.btn_consult.setObjectName("btn_undo")
        self.btn_consult.setToolTip("Copia in workspace per sola lettura (no lock)")
        self.btn_consult.clicked.connect(self._do_consultation)

        self.btn_props_in = QPushButton("⬆  Proprietà → PDM")
        self.btn_props_in.setObjectName("btn_props_in")
        self.btn_props_in.setToolTip("Legge le proprietà custom dal file SW e le salva nel PDM")
        self.btn_props_in.clicked.connect(self._do_import_props)

        self.btn_props_out = QPushButton("⬇  PDM → Proprietà")
        self.btn_props_out.setObjectName("btn_props_out")
        self.btn_props_out.setToolTip("Scrive le proprietà PDM nel file SW")
        self.btn_props_out.clicked.connect(self._do_export_props)

        self.btn_open_app = QPushButton("🚀  Apri PDM")
        self.btn_open_app.setObjectName("btn_open")
        self.btn_open_app.setToolTip("Apre l'applicazione PDM-SW principale")
        self.btn_open_app.clicked.connect(self._do_open_app)

        # 3 colonne × N righe
        grid.addWidget(self.btn_co,        0, 0)
        grid.addWidget(self.btn_ci,        0, 1)
        grid.addWidget(self.btn_undo,      1, 0)
        grid.addWidget(self.btn_consult,   1, 1)
        grid.addWidget(self.btn_props_in,  2, 0)
        grid.addWidget(self.btn_props_out, 2, 1)

        sep = QFrame()
        sep.setObjectName("separator")
        sep.setFrameShape(QFrame.Shape.HLine)
        grid.addWidget(sep, 3, 0, 1, 2)

        grid.addWidget(self.btn_open_app, 4, 0, 1, 2)

        self._main_layout.addWidget(self._grp_actions)

        # --- Genera Codice (visibile solo se doc NON in PDM) ---
        self._grp_generate = QGroupBox()
        gen_layout = QVBoxLayout(self._grp_generate)
        gen_layout.setContentsMargins(6, 6, 6, 6)

        self.btn_generate = QPushButton("➕  Genera Codice PDM")
        self.btn_generate.setObjectName("btn_generate")
        self.btn_generate.setToolTip(
            "Il file non è ancora nel PDM.\nAssegna un codice e crea il documento."
        )
        self.btn_generate.clicked.connect(self._do_generate_code)
        gen_layout.addWidget(self.btn_generate)

        self._main_layout.addWidget(self._grp_generate)
        self._grp_generate.hide()

        # Utente corrente
        uname = self.user.get("full_name", "") if self.user else ""
        lbl_user = QLabel(f"Utente: {uname}")
        lbl_user.setStyleSheet("color:#585b70; font-size:10px;")
        lbl_user.setAlignment(Qt.AlignmentFlag.AlignRight)
        self._main_layout.addWidget(lbl_user)

    # ------------------------------------------------------------------
    def _refresh(self):
        """Ricarica stato documento dal DB e aggiorna la UI."""
        if not self.file_path:
            self._lbl_filename.setText("(nessun file attivo)")
            self._lbl_meta.setText("")
            self._lbl_state.setText("")
            self._lbl_lock.setText("")
            self._set_buttons_enabled(False)
            return

        fp = Path(self.file_path)
        self._lbl_filename.setText(fp.name)

        doc_type = EXT_TO_TYPE.get(fp.suffix.upper(), "?")

        # Ricarica doc dal DB (fresco)
        self.doc = _find_doc(self.db, self.file_path, self.user)

        if not self.doc:
            # File non in PDM
            self._lbl_meta.setText(f"Tipo: {doc_type}  |  Non presente nel PDM")
            self._lbl_state.setText(
                "<span style='color:#fab387;'>⚠ File non registrato nel PDM</span>"
            )
            self._lbl_lock.setText("")
            self._set_buttons_enabled(False)
            self.btn_open_app.setEnabled(True)
            self._grp_generate.show()
            self.adjustSize()
            return

        self._grp_generate.hide()
        d = self.doc
        icon = TYPE_ICONS.get(d["doc_type"], "")
        self._lbl_meta.setText(
            f"{icon} {d['doc_type']}  |  Rev: {d['revision']}  |  {d['code']}"
        )

        state = d["state"]
        color = STATE_COLORS.get(state, "#cdd6f4")
        self._lbl_state.setText(
            f"<span style='background:{color};color:#1e1e2e;"
            f"border-radius:4px;padding:2px 8px;font-weight:bold;'>"
            f"● {state}</span>"
        )

        if d["is_locked"] and self.user:
            locker  = d.get("locked_by_name") or "?"
            is_mine = d["locked_by"] == self.user["id"]
            if is_mine:
                self._lbl_lock.setText("🔒 In checkout da <b>te</b>")
                self._lbl_lock.setStyleSheet("color:#fab387;")
            else:
                ws = d.get("locked_ws", "")
                self._lbl_lock.setText(f"🔒 Bloccato da <b>{locker}</b> [{ws}]")
                self._lbl_lock.setStyleSheet("color:#f38ba8;")
        else:
            self._lbl_lock.setText("🔓 Disponibile")
            self._lbl_lock.setStyleSheet("color:#a6e3a1;")

        self._update_button_states()
        self.adjustSize()

    def _update_button_states(self):
        d = self.doc
        if not d:
            self._set_buttons_enabled(False)
            self.btn_open_app.setEnabled(True)
            return

        state     = d["state"]
        is_locked = bool(d["is_locked"])
        if not self.user:
            return
        is_mine   = is_locked and (d["locked_by"] == self.user["id"])
        readonly  = state in ("Rilasciato", "Obsoleto")

        self.btn_co.setEnabled(not is_locked and not readonly)
        self.btn_ci.setEnabled(is_mine)
        self.btn_undo.setEnabled(is_mine)
        self.btn_consult.setEnabled(bool(d.get("archive_path")))
        self.btn_props_in.setEnabled(is_mine or (not is_locked and not readonly))
        self.btn_props_out.setEnabled(bool(d.get("archive_path")) and is_mine)
        self.btn_open_app.setEnabled(True)

    def _set_buttons_enabled(self, enabled: bool):
        for btn in (self.btn_co, self.btn_ci, self.btn_undo,
                    self.btn_consult, self.btn_props_in, self.btn_props_out):
            btn.setEnabled(enabled)
        self.btn_open_app.setEnabled(True)

    # ------------------------------------------------------------------
    #  Azioni
    # ------------------------------------------------------------------
    def _do_checkout(self):
        if not self.doc or not self.checkout:
            return
        try:
            dest = self.checkout.checkout(self.doc["id"])
            QMessageBox.information(
                self, "Checkout eseguito",
                f"File scaricato in workspace:\n{dest}"
            )
            self._refresh()
        except Exception as e:
            QMessageBox.critical(self, "Errore Checkout", str(e))

    def _do_checkin(self):
        if not self.doc or not self.checkout or not self.props:
            return
        try:
            # Salva il file in SW prima di archiviare
            self._save_active_sw_doc()

            # Import proprietà (silenzioso)
            try:
                props = self.props.read_from_sw_file(Path(self.file_path))
                props.pop("_error", None)
                if props:
                    self.props.save_properties(self.doc["id"], props)
            except Exception as pe:
                logging.warning("Proprietà non importate durante checkin: %s", pe)

            result = self.checkout.checkin(self.doc["id"], archive_file=True)
            msg = "✅  Check-in completato.\n"
            msg += "File archiviato nel PDM.\n" if result.get("archived") else ""
            msg += "⚠ Conflitto rilevato!\n"    if result.get("conflict") else ""
            QMessageBox.information(self, "Check-in", msg)
            self._refresh()
        except Exception as e:
            QMessageBox.critical(self, "Errore Check-in", str(e))

    def _do_undo_checkout(self):
        if not self.doc or not self.checkout:
            return
        r = QMessageBox.question(
            self, "Annulla Checkout",
            "Rilasciare il lock senza archiviare le modifiche?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if r != QMessageBox.StandardButton.Yes:
            return
        try:
            self.checkout.undo_checkout(self.doc["id"])
            QMessageBox.information(self, "Annulla Checkout", "Lock rilasciato.")
            self._refresh()
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

    def _do_consultation(self):
        if not self.doc or not self.checkout:
            return
        try:
            dest = self.checkout.open_for_consultation(self.doc["id"])
            QMessageBox.information(
                self, "Consultazione",
                f"File copiato in workspace (sola lettura):\n{dest}"
            )
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

    def _do_import_props(self):
        if not self.doc or not self.props:
            return
        try:
            fp = Path(self.file_path)
            sync = self.props.sync_sw_to_pdm(
                self.doc["id"], fp, file_name=fp.name
            )
            if not sync.get("ok"):
                QMessageBox.warning(self, "Proprietà", f"Errore lettura SW:\n{sync.get('error', '')}")
                return
            imported = int(sync.get("imported_count", 0))
            updated = bool(sync.get("updated_owner", False))
            if imported <= 0 and not updated:
                QMessageBox.information(
                    self, "Proprietà",
                    "Nessuna proprietà custom trovata nel file SolidWorks."
                )
                return
            msg = f"{imported} proprietà lette da SolidWorks e salvate nel PDM."
            if updated:
                msg += "\nCampi PDM (titolo/descrizione) aggiornati dalla mappatura."
            QMessageBox.information(self, "Proprietà importate", msg)
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

    def _do_export_props(self):
        if not self.doc or not self.props:
            return
        try:
            fp = Path(self.file_path)
            # 1) Sync campi PDM fondamentali (revisione, codice, stato, autore, data)
            sync = self.props.sync_pdm_to_sw(self.doc["id"], fp, force_revision=True)
            if not sync.get("ok"):
                QMessageBox.warning(self, "Proprietà", f"Errore sync PDM->SW:\n{sync.get('error', '')}")
                return
            written = int(sync.get("written_count", 0))
            # 2) Scrivi anche le custom properties dal DB
            custom = self.props.get_properties(self.doc["id"])
            if custom:
                self.props.write_to_sw_file(fp, custom)
            total = written + len(custom)
            if total <= 0:
                QMessageBox.information(
                    self, "Proprietà",
                    "Nessuna proprietà da esportare verso SolidWorks."
                )
                return
            QMessageBox.information(
                self, "Proprietà esportate",
                f"{written} campi PDM + {len(custom)} custom properties scritti nel file SolidWorks."
            )
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

    def _do_open_app(self):
        import subprocess
        script     = ROOT / "main.py"
        venv_py    = ROOT / ".venv" / "Scripts" / "pythonw.exe"
        python_exe = str(venv_py) if venv_py.exists() else sys.executable
        subprocess.Popen([python_exe, str(script)], cwd=str(ROOT))

    def _do_generate_code(self):
        if not self.file_path:
            return
        dlg = CreateCodeDialog(
            self.file_path, self.db, self.sp,
            self.user, self.coding, self.files, parent=self,
        )
        if dlg.exec() == QDialog.DialogCode.Accepted:
            # Aggiorna il percorso del pannello al nuovo file codificato in WS
            if hasattr(dlg, "created_path") and dlg.created_path.exists():
                self.file_path = str(dlg.created_path)
            self._refresh()

    # ------------------------------------------------------------------
    def _save_active_sw_doc(self):
        """Salva il documento attivo in SolidWorks via COM (silenzioso se SW non disponibile)."""
        try:
            import win32com.client
            sw = win32com.client.GetActiveObject("SldWorks.Application")
            if sw:
                doc = sw.ActiveDoc
                if doc:
                    doc.Save()
        except Exception as e:
            logging.debug("Salvataggio SW silenzioso fallito: %s", e)


# ===========================================================================
def main():
    file_path = sys.argv[1] if len(sys.argv) > 1 else ""
    logging.info("pdm_panel avviato  |  file=%s", file_path)

    app = QApplication(sys.argv)
    app.setApplicationName("PDM-SW Panel")
    app.setStyle("Fusion")

    panel = PDMPanel(file_path)
    panel.show()
    panel.raise_()
    panel.activateWindow()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

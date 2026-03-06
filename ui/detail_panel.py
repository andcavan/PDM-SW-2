# =============================================================================
#  ui/detail_panel.py  –  Pannello dettaglio documento (readonly, anteprima)
# =============================================================================
import subprocess
from pathlib import Path

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTabWidget, QTableWidget, QTableWidgetItem, QHeaderView,
    QGroupBox, QFormLayout, QMessageBox, QSizePolicy,
    QStackedWidget, QCheckBox
)
from PyQt6.QtCore import Qt, QSize, pyqtSignal
from PyQt6.QtGui import QPixmap, QFont

from ui.session import session
from ui.styles import STATE_BADGE_STYLE, TYPE_ICON


# Icone fallback grandi per tipo documento (quando manca la thumbnail)
_FALLBACK_ICONS = {
    "Parte":   "🔩",
    "Assieme": "⚙️",
    "Disegno": "📐",
}

_THUMB_SIZE = 200  # px lato thumbnail


class DetailPanel(QWidget):
    """
    Pannello laterale readonly per anteprima rapida di un documento.
    Mostra thumbnail (con fallback icona), info generali, proprietà,
    BOM e storico. Include bottone per apertura in eDrawings.
    Supporta anche la modalità «nodo codice» con pulsanti di creazione.
    """

    # Segnali emessi dai bottoni di creazione (archive-first)
    create_in_sw_requested   = pyqtSignal(int)  # doc_id PRT/ASM senza file
    create_from_file_requested = pyqtSignal(int)  # doc_id PRT/ASM senza file
    add_drw_requested        = pyqtSignal(int)  # parent PRT/ASM doc_id

    def __init__(self, parent=None):
        super().__init__(parent)
        self._current_doc_id: int | None = None
        self._current_code:   str | None = None
        self._code_action_doc_id:    int | None = None  # PRT/ASM senza file
        self._code_prt_asm_for_drw:  int | None = None  # PRT/ASM con file (per DRW)
        self._build_ui()
        self.clear()

    # ------------------------------------------------------------------
    #  Costruzione UI
    # ------------------------------------------------------------------
    def _build_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(0)
        main_layout.setContentsMargins(0, 0, 0, 0)

        self.stack = QStackedWidget()
        main_layout.addWidget(self.stack)

        # ── Pagina 0: modalità documento ──────────────────────────────
        doc_page = QWidget()
        layout = QVBoxLayout(doc_page)
        layout.setSpacing(6)
        layout.setContentsMargins(8, 8, 8, 8)

        # ---- Header: thumbnail + info principali ----
        header = QHBoxLayout()

        # Thumbnail
        self.lbl_thumb = QLabel()
        self.lbl_thumb.setFixedSize(QSize(_THUMB_SIZE, _THUMB_SIZE))
        self.lbl_thumb.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_thumb.setStyleSheet(
            "background-color: #181825; border: 1px solid #313244; border-radius: 6px;"
        )
        header.addWidget(self.lbl_thumb)

        # Info principali (codice, rev, tipo, stato)
        info = QVBoxLayout()
        info.setSpacing(4)

        self.lbl_code = QLabel()
        self.lbl_code.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        self.lbl_code.setStyleSheet("color: #89b4fa;")
        self.lbl_code.setWordWrap(True)
        info.addWidget(self.lbl_code)

        self.lbl_rev = QLabel()
        self.lbl_rev.setStyleSheet("color: #a6adc8; font-size: 12px;")
        info.addWidget(self.lbl_rev)

        self.lbl_type = QLabel()
        self.lbl_type.setStyleSheet("font-size: 13px;")
        info.addWidget(self.lbl_type)

        self.lbl_state = QLabel()
        info.addWidget(self.lbl_state)

        self.lbl_locked = QLabel()
        self.lbl_locked.setStyleSheet("font-size: 11px;")
        info.addWidget(self.lbl_locked)

        info.addStretch()

        # Bottone eDrawings
        self.btn_edrawings = QPushButton("👁️  Apri in eDrawings")
        self.btn_edrawings.setToolTip("Apre il file in eDrawings per consultazione rapida")
        self.btn_edrawings.clicked.connect(self._open_edrawings)
        self.btn_edrawings.setEnabled(False)
        info.addWidget(self.btn_edrawings)

        header.addLayout(info, 1)
        layout.addLayout(header)

        # ---- Tab dettagli (readonly) ----
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)

        self.tabs.addTab(self._build_general_tab(), "Generale")
        self.tabs.addTab(self._build_properties_tab(), "Proprietà SW")
        self.tabs.addTab(self._build_bom_tab(), "Struttura")
        self.tabs.addTab(self._build_history_tab(), "Storico")

        self.stack.addWidget(doc_page)

        # ── Pagina 1: modalità codice ─────────────────────────────────
        self.stack.addWidget(self._build_code_panel())

    # ---- Tab Generale ----
    def _build_general_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setSpacing(8)

        grp = QGroupBox("Informazioni")
        form = QFormLayout(grp)
        form.setSpacing(6)

        self.lbl_title = QLabel("—")
        self.lbl_title.setWordWrap(True)
        form.addRow("Titolo:", self.lbl_title)

        self.lbl_desc = QLabel("—")
        self.lbl_desc.setWordWrap(True)
        self.lbl_desc.setStyleSheet("color: #a6adc8;")
        form.addRow("Descrizione:", self.lbl_desc)

        self.lbl_author = QLabel("—")
        form.addRow("Creato da:", self.lbl_author)

        self.lbl_date = QLabel("—")
        form.addRow("Data mod.:", self.lbl_date)

        self.lbl_file = QLabel("—")
        self.lbl_file.setWordWrap(True)
        self.lbl_file.setStyleSheet("color: #a6adc8; font-size: 11px;")
        form.addRow("File:", self.lbl_file)

        layout.addWidget(grp)
        layout.addStretch()
        return w

    # ---- Tab Proprietà SW ----
    def _build_properties_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        self.tbl_props = QTableWidget(0, 2)
        self.tbl_props.setHorizontalHeaderLabels(["Proprietà", "Valore"])
        self.tbl_props.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeMode.Stretch
        )
        self.tbl_props.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        layout.addWidget(self.tbl_props)
        return w

    # ---- Tab BOM ----
    def _build_bom_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        self.tbl_bom = QTableWidget(0, 5)
        self.tbl_bom.setHorizontalHeaderLabels(
            ["Codice", "Rev.", "Tipo", "Titolo", "Qtà"]
        )
        self.tbl_bom.horizontalHeader().setSectionResizeMode(
            3, QHeaderView.ResizeMode.Stretch
        )
        self.tbl_bom.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        layout.addWidget(self.tbl_bom)
        return w

    # ---- Tab Storico ----
    def _build_history_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        self.tbl_history = QTableWidget(0, 5)
        self.tbl_history.setHorizontalHeaderLabels(
            ["Data/Ora", "Azione", "Da", "A", "Utente"]
        )
        self.tbl_history.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch
        )
        self.tbl_history.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        layout.addWidget(self.tbl_history)
        return w

    # ------------------------------------------------------------------
    #  API pubblica
    # ------------------------------------------------------------------
    def clear(self):
        """Svuota il pannello (nessun documento selezionato)."""
        self._current_doc_id = None
        self._current_code   = None
        self.stack.setCurrentIndex(0)
        self.lbl_thumb.setText("📄")
        self.lbl_thumb.setStyleSheet(
            "background-color: #181825; border: 1px solid #313244; "
            "border-radius: 6px; font-size: 48px; color: #45475a;"
        )
        self.lbl_code.setText("Seleziona un documento")
        self.lbl_rev.setText("")
        self.lbl_type.setText("")
        self.lbl_state.setText("")
        self.lbl_locked.setText("")
        self.lbl_title.setText("—")
        self.lbl_desc.setText("—")
        self.lbl_author.setText("—")
        self.lbl_date.setText("—")
        self.lbl_file.setText("—")
        self.btn_edrawings.setEnabled(False)
        self.tbl_props.setRowCount(0)
        self.tbl_bom.setRowCount(0)
        self.tbl_history.setRowCount(0)

    def load_document(self, doc_id: int):
        """Carica e mostra i dettagli del documento indicato."""
        if not session.is_connected or not doc_id:
            self.clear()
            return

        doc = session.files.get_document(doc_id)
        if not doc:
            self.clear()
            return

        self._current_doc_id = doc_id
        self.stack.setCurrentIndex(0)

        # ---- Thumbnail ----
        self._load_thumbnail(doc)

        # ---- Header ----
        self.lbl_code.setText(doc["code"])
        self.lbl_rev.setText(f"Revisione {doc['revision']}")
        icon = TYPE_ICON.get(doc["doc_type"], "")
        self.lbl_type.setText(f"{icon}  {doc['doc_type']}")

        state = doc["state"]
        badge = STATE_BADGE_STYLE.get(state, "")
        self.lbl_state.setText(f"<span style='{badge}'>{state}</span>")
        self.lbl_state.setTextFormat(Qt.TextFormat.RichText)

        if doc["is_locked"]:
            locked_name = doc.get("locked_by_name", "sconosciuto")
            self.lbl_locked.setText(f"🔒 In checkout: {locked_name}")
            self.lbl_locked.setStyleSheet("color:#fab387; font-size: 11px;")
        else:
            self.lbl_locked.setText("✅ Disponibile")
            self.lbl_locked.setStyleSheet("color:#a6e3a1; font-size: 11px;")

        # ---- Tab Generale ----
        self.lbl_title.setText(doc.get("title") or "—")
        self.lbl_desc.setText(doc.get("description") or "—")
        self.lbl_author.setText(doc.get("created_by_name") or "—")
        mod_at = doc.get("modified_at") or doc.get("created_at") or ""
        self.lbl_date.setText(str(mod_at)[:19] if mod_at else "—")

        if doc.get("file_name"):
            file_info = doc["file_name"]
            if doc.get("archive_path"):
                file_info += f"\n[{doc['archive_path']}]"
            else:
                file_info += "\n(non archiviato)"
            self.lbl_file.setText(file_info)
        else:
            self.lbl_file.setText("Nessun file associato")

        # eDrawings: abilitato solo se c'è un file archiviato
        self.btn_edrawings.setEnabled(bool(doc.get("archive_path")))

        # ---- Tab Proprietà ----
        self.tbl_props.setRowCount(0)
        props = session.properties.get_properties(doc_id)
        for name, value in props.items():
            row = self.tbl_props.rowCount()
            self.tbl_props.insertRow(row)
            self.tbl_props.setItem(row, 0, QTableWidgetItem(name))
            self.tbl_props.setItem(row, 1, QTableWidgetItem(value))

        # ---- Tab BOM ----
        self.tbl_bom.setRowCount(0)
        comps = session.asm.get_components(doc_id)
        for c in comps:
            row = self.tbl_bom.rowCount()
            self.tbl_bom.insertRow(row)
            self.tbl_bom.setItem(row, 0, QTableWidgetItem(c["code"]))
            self.tbl_bom.setItem(row, 1, QTableWidgetItem(c["revision"]))
            self.tbl_bom.setItem(
                row, 2, QTableWidgetItem(
                    TYPE_ICON.get(c["doc_type"], "") + " " + c["doc_type"]
                )
            )
            self.tbl_bom.setItem(row, 3, QTableWidgetItem(c["title"]))
            self.tbl_bom.setItem(row, 4, QTableWidgetItem(str(c["quantity"])))

        # ---- Tab Storico ----
        self.tbl_history.setRowCount(0)
        rows = session.workflow.get_history(doc_id)
        if session.checkout:
            rows += session.checkout.get_log(doc_id)
        rows.sort(
            key=lambda r: r.get("changed_at") or r.get("timestamp", ""),
            reverse=True,
        )
        for r in rows:
            i = self.tbl_history.rowCount()
            self.tbl_history.insertRow(i)
            ts = r.get("changed_at") or r.get("timestamp", "")
            self.tbl_history.setItem(i, 0, QTableWidgetItem(str(ts)[:19]))
            self.tbl_history.setItem(i, 1, QTableWidgetItem(r.get("action", "cambio stato")))
            self.tbl_history.setItem(i, 2, QTableWidgetItem(r.get("from_state", "")))
            self.tbl_history.setItem(i, 3, QTableWidgetItem(r.get("to_state", "")))
            self.tbl_history.setItem(i, 4, QTableWidgetItem(r.get("user_name", "")))

    # ------------------------------------------------------------------
    #  Pannello modalità codice
    # ------------------------------------------------------------------
    def _build_code_panel(self) -> QWidget:
        """Pannello visualizzato quando è selezionato un nodo codice."""
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setSpacing(10)
        layout.setContentsMargins(8, 8, 8, 8)

        # Header: icona + codice + titolo
        hdr = QHBoxLayout()
        lbl_icon = QLabel("📁")
        lbl_icon.setStyleSheet("font-size: 40px;")
        lbl_icon.setFixedWidth(56)
        hdr.addWidget(lbl_icon)

        code_info = QVBoxLayout()
        self.lbl_code_str = QLabel()
        self.lbl_code_str.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        self.lbl_code_str.setStyleSheet("color: #89b4fa;")
        code_info.addWidget(self.lbl_code_str)

        self.lbl_code_title = QLabel()
        self.lbl_code_title.setStyleSheet("color: #a6adc8; font-size: 12px;")
        self.lbl_code_title.setWordWrap(True)
        code_info.addWidget(self.lbl_code_title)
        code_info.addStretch()
        hdr.addLayout(code_info, 1)
        layout.addLayout(hdr)

        # File in archivio
        grp_files = QGroupBox("File in archivio")
        files_layout = QVBoxLayout(grp_files)
        self.lbl_prt_status = QLabel()
        self.lbl_asm_status = QLabel()
        self.lbl_drw_status = QLabel()
        for lbl in (self.lbl_prt_status, self.lbl_asm_status, self.lbl_drw_status):
            lbl.setStyleSheet("font-size: 13px; padding: 2px;")
            files_layout.addWidget(lbl)
        layout.addWidget(grp_files)

        # Azioni
        self.grp_actions = QGroupBox("Azioni")
        act_layout = QVBoxLayout(self.grp_actions)

        self.btn_create_sw = QPushButton("🔨  Crea in SW")
        self.btn_create_sw.setObjectName("btn_primary")
        self.btn_create_sw.clicked.connect(self._on_create_in_sw)
        act_layout.addWidget(self.btn_create_sw)

        self.btn_create_file = QPushButton("📂  Crea da file")
        self.btn_create_file.clicked.connect(self._on_create_from_file)
        act_layout.addWidget(self.btn_create_file)

        self.btn_add_drw = QPushButton("📐  Aggiungi DRW")
        self.btn_add_drw.clicked.connect(self._on_add_drw)
        act_layout.addWidget(self.btn_add_drw)

        self.chk_checkout = QCheckBox("Metti in checkout dopo la creazione")
        act_layout.addWidget(self.chk_checkout)

        layout.addWidget(self.grp_actions)
        layout.addStretch()
        return w

    def load_code(self, code: str, docs: list[dict]):
        """Carica e mostra le informazioni del nodo codice (modalità codice)."""
        self._current_doc_id = None
        self._current_code = code
        self.stack.setCurrentIndex(1)

        # Header
        self.lbl_code_str.setText(code)
        rep = max(docs, key=lambda d: d["revision"]) if docs else None
        self.lbl_code_title.setText(rep.get("title", "") if rep else "")

        # Status per tipo
        prt = next((d for d in docs if d["doc_type"] == "Parte"),   None)
        asm = next((d for d in docs if d["doc_type"] == "Assieme"), None)
        drw = next((d for d in docs if d["doc_type"] == "Disegno"), None)

        def _status_text(doc) -> tuple[str, str]:
            if not doc:
                return "", ""
            icon = TYPE_ICON.get(doc["doc_type"], "")
            rev  = doc["revision"]
            if doc.get("archive_path"):
                return (f"{icon}  {doc['doc_type']}  Rev.{rev}  ✅",
                        "color: #a6e3a1;")
            if doc["is_locked"]:
                return (f"{icon}  {doc['doc_type']}  Rev.{rev}  🔒 In checkout",
                        "color: #fab387;")
            return (f"{icon}  {doc['doc_type']}  Rev.{rev}  ❌ Non archiviato",
                    "color: #f38ba8;")

        for doc, lbl in (
            (prt, self.lbl_prt_status),
            (asm, self.lbl_asm_status),
            (drw, self.lbl_drw_status),
        ):
            text, style = _status_text(doc)
            lbl.setText(text)
            lbl.setStyleSheet(f"font-size: 13px; padding: 2px; {style}")
            lbl.setVisible(bool(text))

        # Calcola candidati azioni
        prt_asm_no_file = [
            d for d in docs
            if d["doc_type"] in ("Parte", "Assieme")
            and not d.get("archive_path") and not d["is_locked"]
        ]
        prt_asm_with_file = [
            d for d in docs
            if d["doc_type"] in ("Parte", "Assieme") and d.get("archive_path")
        ]
        drw_with_file = [
            d for d in docs
            if d["doc_type"] == "Disegno" and d.get("archive_path")
        ]

        self._code_action_doc_id   = prt_asm_no_file[0]["id"]   if prt_asm_no_file   else None
        self._code_prt_asm_for_drw = prt_asm_with_file[0]["id"] if prt_asm_with_file else None

        need_drw = bool(prt_asm_with_file) and not bool(drw_with_file)
        self.btn_create_sw.setVisible(bool(prt_asm_no_file))
        self.btn_create_file.setVisible(bool(prt_asm_no_file))
        self.btn_add_drw.setVisible(need_drw)
        self.grp_actions.setVisible(bool(prt_asm_no_file) or need_drw)

    def _on_create_in_sw(self):
        if self._code_action_doc_id:
            self.create_in_sw_requested.emit(self._code_action_doc_id)

    def _on_create_from_file(self):
        if self._code_action_doc_id:
            self.create_from_file_requested.emit(self._code_action_doc_id)

    def _on_add_drw(self):
        if self._code_prt_asm_for_drw:
            self.add_drw_requested.emit(self._code_prt_asm_for_drw)

    # ------------------------------------------------------------------
    #  Thumbnail
    # ------------------------------------------------------------------
    def _load_thumbnail(self, doc: dict):
        """Carica thumbnail da file o mostra icona fallback."""
        thumb_path = self._get_thumbnail_path(doc)

        if thumb_path and thumb_path.exists():
            pixmap = QPixmap(str(thumb_path))
            if not pixmap.isNull():
                scaled = pixmap.scaled(
                    _THUMB_SIZE, _THUMB_SIZE,
                    Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation,
                )
                self.lbl_thumb.setPixmap(scaled)
                self.lbl_thumb.setStyleSheet(
                    "background-color: #181825; border: 1px solid #313244; "
                    "border-radius: 6px;"
                )
                return

        # Fallback: icona tipo grande
        icon = _FALLBACK_ICONS.get(doc["doc_type"], "📄")
        self.lbl_thumb.setText(icon)
        self.lbl_thumb.setPixmap(QPixmap())  # Pulisci eventuale pixmap
        self.lbl_thumb.setText(icon)
        self.lbl_thumb.setStyleSheet(
            "background-color: #181825; border: 1px solid #313244; "
            "border-radius: 6px; font-size: 64px; color: #585b70;"
        )

    def _get_thumbnail_path(self, doc: dict) -> "Path | None":
        """Restituisce il path della thumbnail se configurato."""
        if not session.sp:
            return None
        code = doc.get("code", "")
        rev = doc.get("revision", "")
        if not code:
            return None
        thumb_file = session.sp.thumbnails / f"{code}_{rev}.png"
        return thumb_file

    # ------------------------------------------------------------------
    #  eDrawings
    # ------------------------------------------------------------------
    def _open_edrawings(self):
        """Apre il file archiviato in eDrawings per consultazione."""
        if not self._current_doc_id:
            return

        doc = session.files.get_document(self._current_doc_id)
        if not doc or not doc.get("archive_path"):
            QMessageBox.warning(
                self, "File non disponibile",
                "Nessun file archiviato per questo documento."
            )
            return

        # Costruisci path completo archivo
        archive_file = session.sp.root / doc["archive_path"]
        if not archive_file.exists():
            QMessageBox.warning(
                self, "File non trovato",
                f"Il file non è stato trovato nell'archivio:\n{archive_file}"
            )
            return

        try:
            from ui.sw_config_dialog import SWConfigDialog
            edrawings_exe = SWConfigDialog.get_edrawings_exe()

            if edrawings_exe:
                subprocess.Popen([str(edrawings_exe), str(archive_file)])
            else:
                QMessageBox.warning(
                    self, "eDrawings non configurato",
                    "Eseguibile eDrawings non configurato.\n\n"
                    "Aprire Strumenti → Configurazione SolidWorks → tab SolidWorks\n"
                    "e impostare il percorso di eDrawings (o usare 'Rileva')."
                )
        except Exception as e:
            QMessageBox.critical(
                self, "Errore",
                f"Impossibile aprire eDrawings:\n{e}"
            )

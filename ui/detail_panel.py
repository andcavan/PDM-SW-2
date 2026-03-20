# =============================================================================
#  ui/detail_panel.py  –  Pannello dettaglio documento (readonly, anteprima)
# =============================================================================
import shutil
import subprocess
from datetime import date
from pathlib import Path

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTabWidget, QTableWidget, QTableWidgetItem, QHeaderView,
    QGroupBox, QFormLayout, QMessageBox, QSizePolicy,
    QStackedWidget, QTextEdit, QToolBar
)
from PyQt6.QtCore import Qt, QSize, pyqtSignal
from PyQt6.QtGui import QPixmap, QFont, QAction, QKeySequence

from ui.session import session
from ui.styles import STATE_BADGE_STYLE, TYPE_ICON


# Icone fallback grandi per tipo documento (quando manca la thumbnail)
_FALLBACK_ICONS = {
    "Parte":   "🔩",
    "Assieme": "⚙️",
    "Disegno": "📐",
}

_THUMB_W = 360   # px larghezza thumbnail principale
_THUMB_H = 240   # px altezza thumbnail principale (rapporto 1.5:1)
_SUB_W   = 300   # px larghezza sub-thumbnail (code-node)
_SUB_H   = 200   # px altezza sub-thumbnail (rapporto 1.5:1)

_EXT_FOR_TYPE = {
    "Parte":   ".SLDPRT",
    "Assieme": ".SLDASM",
    "Disegno": ".SLDDRW",
}


class DetailPanel(QWidget):
    """
    Pannello laterale readonly per anteprima rapida di un documento.
    Mostra thumbnail (con fallback icona), info generali, proprietà,
    BOM e storico. Include bottone per apertura in eDrawings.
    Supporta anche la modalità «nodo codice» con pulsanti di creazione.
    """

    # Segnali emessi dai bottoni di creazione (archive-first)
    create_in_sw_requested        = pyqtSignal(int)  # doc_id PRT/ASM senza file
    create_from_file_requested    = pyqtSignal(int)  # doc_id PRT/ASM senza file
    add_drw_requested             = pyqtSignal(int)  # parent PRT/ASM doc_id
    create_drw_from_file_requested = pyqtSignal(int)  # parent PRT/ASM doc_id
    # Segnale emesso dopo ogni azione workflow (per aggiornare l'albero)
    workflow_action_completed     = pyqtSignal()
    # Segnale emesso per delegare l'apertura del dialog workflow all'ArchiveView
    # args: rep (dict), all_docs (list), target_state (str — "" = nessuna pre-selezione)
    workflow_requested            = pyqtSignal(dict, list, str)
    # Segnale interno per aggiornamento UI da thread PDF (thread-safe)
    _pdf_done = pyqtSignal(bool, str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._current_doc_id: int | None = None
        self._current_code:   str | None = None
        self._code_action_doc_id:    int | None = None  # PRT/ASM senza file
        self._code_prt_asm_for_drw:  int | None = None  # PRT/ASM con file (per DRW)
        self._code_prt_doc:  dict | None = None
        self._code_drw_doc:  dict | None = None
        self._code_docs: list[dict] = []
        self._code_representative: dict | None = None
        self._build_ui()
        self._pdf_done.connect(self._on_pdf_done)
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
        self.lbl_thumb.setFixedSize(QSize(_THUMB_W, _THUMB_H))
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

        # Bottone toggle Note
        self.btn_toggle_notes = QPushButton("📝  Note")
        self.btn_toggle_notes.setToolTip("Mostra/nascondi il tab Note")
        self.btn_toggle_notes.setCheckable(True)
        self.btn_toggle_notes.setChecked(False)
        self.btn_toggle_notes.clicked.connect(self._on_toggle_notes)
        info.addWidget(self.btn_toggle_notes)

        header.addLayout(info, 1)
        layout.addLayout(header)

        # ---- Tab dettagli (readonly) ----
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)

        self.tabs.addTab(self._build_general_tab(), "Generale")
        self.tabs.addTab(self._build_properties_tab(), "Proprietà SW")
        self._bom_tab_index = self.tabs.count()
        self.tabs.addTab(self._build_bom_tab(), "Struttura")
        self.tabs.addTab(self._build_history_tab(), "Storico")
        self._revisions_tab_index = self.tabs.count()
        self.tabs.addTab(self._build_revisions_tab(), "Revisioni")
        self.tabs.setTabVisible(self._revisions_tab_index, False)
        self._notes_tab_index = self.tabs.count()
        self.tabs.addTab(self._build_notes_tab(doc_mode=True), "Note")
        self.tabs.setTabVisible(self._notes_tab_index, False)

        # ---- Azioni DRW (visibile solo per PRT/ASM senza disegno archiviato) ----
        self.grp_doc_drw = QGroupBox("Disegno")
        doc_drw_row = QHBoxLayout(self.grp_doc_drw)
        doc_drw_row.setSpacing(6)
        self.btn_doc_add_drw = QPushButton("📐  Aggiungi DRW")
        self.btn_doc_add_drw.setToolTip("Crea il disegno da template SolidWorks")
        self.btn_doc_add_drw.clicked.connect(self._on_doc_add_drw)
        doc_drw_row.addWidget(self.btn_doc_add_drw)
        self.btn_doc_create_drw_from_file = QPushButton("📂  Crea DRW da file")
        self.btn_doc_create_drw_from_file.setToolTip("Importa un disegno esistente via SolidWorks")
        self.btn_doc_create_drw_from_file.clicked.connect(self._on_doc_create_drw_from_file)
        doc_drw_row.addWidget(self.btn_doc_create_drw_from_file)
        doc_drw_row.addStretch()
        self.grp_doc_drw.setVisible(False)
        layout.addWidget(self.grp_doc_drw)

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

        # ---- PDF ----
        self.grp_pdf = QGroupBox("PDF")
        pdf_layout = QVBoxLayout(self.grp_pdf)
        pdf_layout.setSpacing(4)

        self.lbl_pdf_path = QLabel("—")
        self.lbl_pdf_path.setWordWrap(True)
        self.lbl_pdf_path.setStyleSheet("color: #a6adc8; font-size: 11px;")
        pdf_layout.addWidget(self.lbl_pdf_path)

        pdf_btns = QHBoxLayout()
        self.btn_gen_pdf = QPushButton("🔄  Genera PDF")
        self.btn_gen_pdf.setToolTip("Genera PDF del disegno (solo per Disegno archiviato)")
        self.btn_gen_pdf.clicked.connect(self._on_generate_pdf)
        pdf_btns.addWidget(self.btn_gen_pdf)

        self.btn_open_pdf = QPushButton("👁️  Apri PDF")
        self.btn_open_pdf.setToolTip("Apre il PDF con il visualizzatore predefinito")
        self.btn_open_pdf.clicked.connect(self._on_open_pdf)
        self.btn_open_pdf.setEnabled(False)
        pdf_btns.addWidget(self.btn_open_pdf)

        self.btn_save_pdf = QPushButton("💾  Salva PDF")
        self.btn_save_pdf.setToolTip("Salva una copia del PDF in una cartella a scelta")
        self.btn_save_pdf.clicked.connect(self._on_save_pdf)
        self.btn_save_pdf.setEnabled(False)
        pdf_btns.addWidget(self.btn_save_pdf)

        pdf_layout.addLayout(pdf_btns)

        self.grp_pdf.setVisible(False)
        layout.addWidget(self.grp_pdf)

        layout.addStretch()
        return w

    # ---- Tab Proprietà SW ----
    def _build_properties_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)

        btn_row = QHBoxLayout()
        self.btn_sync_props = QPushButton("🔄  Aggiorna Proprietà")
        self.btn_sync_props.setToolTip(
            "Sincronizza PDM ↔ SW secondo la mappatura e importa tutte le proprietà SW nel PDM"
        )
        self.btn_sync_props.clicked.connect(self._on_sync_props)
        self.btn_import_props = QPushButton("⬇️  Importa da SW")
        self.btn_import_props.setToolTip(
            "Legge le proprietà custom dal file SolidWorks archiviato\n"
            "(non richiede SolidWorks aperto)"
        )
        self.btn_import_props.clicked.connect(self._on_import_from_sw)
        self.btn_export_props = QPushButton("⬆️  Esporta in SW")
        self.btn_export_props.setToolTip("Esporta le proprietà PDM nel file SW (segue la mappatura)")
        self.btn_export_props.clicked.connect(self._on_export_to_sw)
        for b in (self.btn_sync_props, self.btn_import_props, self.btn_export_props):
            btn_row.addWidget(b)
        btn_row.addStretch()
        layout.addLayout(btn_row)

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

    # ---- Tab Revisioni ----
    def _build_revisions_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)

        # Barra pulsanti apertura file revisione obsoleta (in cima, come tab Proprietà)
        btn_bar = QHBoxLayout()
        self.btn_rev_edrawings = QPushButton("👁️ Apri in eDrawings")
        self.btn_rev_sw        = QPushButton("🔧 Apri in SW")
        self.btn_rev_pdf       = QPushButton("📄 Apri PDF")
        for b in (self.btn_rev_edrawings, self.btn_rev_sw, self.btn_rev_pdf):
            b.setEnabled(False)
            btn_bar.addWidget(b)
        btn_bar.addStretch()
        layout.addLayout(btn_bar)

        self.tbl_revisions = QTableWidget(0, 4)
        self.tbl_revisions.setHorizontalHeaderLabels(
            ["Revisione", "Data rilascio", "Data sostituzione", "Note"]
        )
        hdr = self.tbl_revisions.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self.tbl_revisions.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tbl_revisions.setAlternatingRowColors(True)
        self.tbl_revisions.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        layout.addWidget(self.tbl_revisions)

        self.tbl_revisions.itemSelectionChanged.connect(self._on_revision_selection_changed)
        self.btn_rev_edrawings.clicked.connect(lambda: self._open_obsolete("edrawings"))
        self.btn_rev_sw.clicked.connect(lambda: self._open_obsolete("sw"))
        self.btn_rev_pdf.clicked.connect(lambda: self._open_obsolete("pdf"))
        return w

    # ---- Tab Note (doc mode) / GroupBox Note (code mode) ----
    def _build_notes_tab(self, doc_mode: bool = True) -> QWidget:
        """
        Costruisce il pannello Note (rich text).
        doc_mode=True → usato come tab in self.tabs (doc mode).
        doc_mode=False → usato come GroupBox in code panel.
        """
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setSpacing(4)
        layout.setContentsMargins(4, 4, 4, 4)

        # Toolbar rich text
        toolbar = QToolBar()
        toolbar.setIconSize(QSize(16, 16))

        if doc_mode:
            self._notes_edit = QTextEdit()
            edit = self._notes_edit
        else:
            self._notes_edit_code = QTextEdit()
            edit = self._notes_edit_code

        edit.setAcceptRichText(True)
        edit.setMinimumHeight(120)

        act_bold = QAction("B", w)
        act_bold.setCheckable(True)
        act_bold.setShortcut(QKeySequence.StandardKey.Bold)
        act_bold.triggered.connect(
            lambda checked: edit.setFontWeight(
                700 if checked else 400
            )
        )
        act_bold.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
        toolbar.addAction(act_bold)

        act_italic = QAction("I", w)
        act_italic.setCheckable(True)
        act_italic.setShortcut(QKeySequence.StandardKey.Italic)
        act_italic.triggered.connect(edit.setFontItalic)
        f = QFont("Segoe UI", 9)
        f.setItalic(True)
        act_italic.setFont(f)
        toolbar.addAction(act_italic)

        act_underline = QAction("U", w)
        act_underline.setCheckable(True)
        act_underline.setShortcut(QKeySequence.StandardKey.Underline)
        act_underline.triggered.connect(edit.setFontUnderline)
        f2 = QFont("Segoe UI", 9)
        f2.setUnderline(True)
        act_underline.setFont(f2)
        toolbar.addAction(act_underline)

        toolbar.addSeparator()

        act_date = QAction("📅 Data", w)
        act_date.setToolTip("Inserisce data odierna e nome utente")
        act_date.triggered.connect(lambda _checked=False: self._insert_date_user(edit))
        toolbar.addAction(act_date)

        layout.addWidget(toolbar)
        layout.addWidget(edit, 1)

        # Footer: save button + last-updated label
        footer = QHBoxLayout()
        btn_save = QPushButton("💾  Salva")
        if doc_mode:
            self._notes_lbl_updated = QLabel("")
            btn_save.clicked.connect(self._save_notes_doc)
            lbl_updated = self._notes_lbl_updated
        else:
            self._notes_lbl_updated_code = QLabel("")
            btn_save.clicked.connect(self._save_notes_code)
            lbl_updated = self._notes_lbl_updated_code

        lbl_updated.setStyleSheet("color: #a6adc8; font-size: 11px;")
        footer.addWidget(btn_save)
        footer.addWidget(lbl_updated, 1)
        layout.addLayout(footer)

        return w

    # ------------------------------------------------------------------
    #  Helpers statici
    # ------------------------------------------------------------------
    @staticmethod
    def _get_code_representative_doc(docs: list[dict]) -> "dict | None":
        """
        Restituisce il documento 'rappresentativo' del codice per il workflow.
        Priorità: PRT/ASM attivo (rev. più alta) → DRW attivo → PRT/ASM obsoleto → qualsiasi.
        """
        def latest(pool):
            return max(pool, key=lambda d: d["revision"]) if pool else None

        active_prt_asm = [d for d in docs
                          if d["doc_type"] in ("Parte", "Assieme")
                          and d["state"] != "Obsoleto"]
        if active_prt_asm:
            return latest(active_prt_asm)

        active_drw = [d for d in docs
                      if d["doc_type"] == "Disegno"
                      and d["state"] != "Obsoleto"]
        if active_drw:
            return latest(active_drw)

        all_prt_asm = [d for d in docs if d["doc_type"] in ("Parte", "Assieme")]
        return latest(all_prt_asm) or latest(docs)

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
        self.grp_doc_drw.setVisible(False)
        self.grp_pdf.setVisible(False)
        self.tbl_props.setRowCount(0)
        self.tbl_bom.setRowCount(0)
        self.tbl_history.setRowCount(0)
        self.tbl_revisions.setRowCount(0)
        self.tabs.setTabVisible(self._revisions_tab_index, False)
        # Nascondi note e resetta toggle
        self.tabs.setTabVisible(self._notes_tab_index, False)
        self.btn_toggle_notes.setChecked(False)
        self._notes_edit.setPlainText("")
        self._notes_lbl_updated.setText("")
        self.grp_notes_code.setVisible(False)
        self.btn_toggle_notes_code.setChecked(False)
        self._notes_edit_code.setPlainText("")
        self._notes_lbl_updated_code.setText("")
        # Reset sezione workflow codice
        self._code_docs = []
        self._code_representative = None
        self.lbl_code_wf_state.setText("")
        self.lbl_code_wf_rev.setText("")
        for _b in (self.btn_wf_release, self.btn_wf_obsolete,
                   self.btn_wf_new_rev, self.btn_wf_cancel_rev):
            _b.setVisible(False)

    def load_document(self, doc_id: int):
        """Carica e mostra i dettagli del documento indicato."""
        if not session.is_connected or not doc_id:
            self.clear()
            return

        doc = session.files.get_document(doc_id)
        if not doc:
            self.clear()
            return

        _same_doc = (doc_id == self._current_doc_id)
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

        # ---- PDF (solo per Disegno archiviato) ----
        is_drw_archived = (
            doc.get("doc_type") == "Disegno" and bool(doc.get("archive_path"))
        )
        self.grp_pdf.setVisible(is_drw_archived)
        if is_drw_archived:
            is_obsolete_drw = (state == "Obsoleto")
            self.grp_pdf.setTitle("PDF  ⚠️ revisione obsoleta" if is_obsolete_drw else "PDF")
            pdf_path = doc.get("pdf_path") or ""
            has_pdf = False
            if pdf_path:
                # Support both relative paths (new) and absolute paths (legacy)
                p = (session.sp.root / pdf_path) if session.sp else Path(pdf_path)
                if not p.exists() and Path(pdf_path).is_absolute():
                    p = Path(pdf_path)
                has_pdf = p.exists()
            if has_pdf:
                self.lbl_pdf_path.setText(pdf_path)
            elif pdf_path:
                self.lbl_pdf_path.setText(f"File non trovato:\n{pdf_path}")
            else:
                self.lbl_pdf_path.setText("PDF non ancora generato")
            # Generazione PDF disabilitata per stati definitivi
            can_gen_pdf = state not in ("Rilasciato", "Obsoleto")
            self.btn_gen_pdf.setEnabled(can_gen_pdf)
            self.btn_gen_pdf.setToolTip(
                "Il PDF è stato generato al rilascio e non può essere rigenerato"
                if not can_gen_pdf else
                "Genera PDF del disegno"
            )
            self.btn_open_pdf.setEnabled(has_pdf)
            self.btn_save_pdf.setEnabled(has_pdf)

        # ---- Note ----
        self._current_code = doc.get("code", "")
        if not _same_doc:
            self._load_notes_to(self._notes_edit, self._notes_lbl_updated, self._current_code)
        is_obsoleto = doc.get("state") == "Obsoleto"
        self._notes_edit.setReadOnly(is_obsoleto)

        # ---- Azioni DRW (visibile per PRT/ASM senza disegno archiviato) ----
        if doc["doc_type"] in ("Parte", "Assieme") and doc.get("archive_path"):
            drw_doc = session.db.fetchone(
                "SELECT id, archive_path, is_locked FROM documents "
                "WHERE code=? AND doc_type='Disegno' AND state != 'Obsoleto' "
                "ORDER BY revision DESC",
                (doc["code"],),
            ) if session.db else None
            drw_archived = bool(drw_doc and drw_doc.get("archive_path"))
            drw_locked = bool(drw_doc and drw_doc["is_locked"])
            show_drw_btns = not drw_archived and not drw_locked
            self.grp_doc_drw.setVisible(show_drw_btns)
        else:
            self.grp_doc_drw.setVisible(False)

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
        def _collect(parent_id: int, depth: int = 0, visited=None):
            if visited is None:
                visited = set()
            if parent_id in visited:
                return []
            visited.add(parent_id)
            rows = []
            for comp in session.asm.get_components(parent_id):
                rows.append((depth, comp))
                if comp.get("doc_type") == "Assieme":
                    rows.extend(_collect(comp["child_id"], depth + 1, visited))
            return rows

        comps = _collect(doc_id)
        for depth, c in comps:
            row = self.tbl_bom.rowCount()
            self.tbl_bom.insertRow(row)
            indent = ("  " * depth) + ("└─ " if depth > 0 else "")
            self.tbl_bom.setItem(row, 0, QTableWidgetItem(indent + c["code"]))
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

        # ---- Tab Revisioni (revisioni obsolete dello stesso codice+tipo) ----
        self.tbl_revisions.setRowCount(0)
        if session.files:
            code     = doc.get("code", "")
            doc_type = doc.get("doc_type", "")
            prev_docs = [
                d for d in session.files.search_documents(code=code)
                if d["code"] == code
                and d["doc_type"] == doc_type
                and d["state"] == "Obsoleto"
            ]
            prev_docs.sort(key=lambda d: d["revision"], reverse=True)
            for pd in prev_docs:
                hist = session.workflow.get_history(pd["id"])
                release_at = next(
                    (h["changed_at"] for h in sorted(hist, key=lambda h: h["changed_at"])
                     if h.get("to_state") == "Rilasciato"), "—"
                )
                obsolete_at = next(
                    (h["changed_at"] for h in sorted(hist, key=lambda h: h["changed_at"])
                     if h.get("to_state") == "Obsoleto"), "—"
                )
                notes_parts = [h["notes"] for h in hist
                               if h.get("to_state") == "Obsoleto" and h.get("notes")]
                note = notes_parts[-1] if notes_parts else ""
                i = self.tbl_revisions.rowCount()
                self.tbl_revisions.insertRow(i)
                item_rev = QTableWidgetItem(pd["revision"])
                item_rev.setData(Qt.ItemDataRole.UserRole, pd["id"])
                self.tbl_revisions.setItem(i, 0, item_rev)
                self.tbl_revisions.setItem(i, 1, QTableWidgetItem(str(release_at)[:19]))
                self.tbl_revisions.setItem(i, 2, QTableWidgetItem(str(obsolete_at)[:19]))
                self.tbl_revisions.setItem(i, 3, QTableWidgetItem(note))
            has_prev = bool(prev_docs)
            self.tabs.setTabVisible(self._revisions_tab_index, has_prev)
            if has_prev:
                self.tabs.setTabText(self._revisions_tab_index,
                                     f"Revisioni ({len(prev_docs)})")

        # ---- Tab Struttura: visibile solo per Assieme ----
        self.tabs.setTabVisible(self._bom_tab_index, doc.get("doc_type") == "Assieme")

        # ---- Bottoni Proprietà SW: disabilitati per Rilasciato/Obsoleto (tranne Admin) ----
        is_admin = session.can("admin")
        sw_locked = state in ("Rilasciato", "Obsoleto") and not is_admin
        for btn in (self.btn_sync_props, self.btn_import_props, self.btn_export_props):
            btn.setEnabled(not sw_locked)

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

        self.btn_toggle_notes_code = QPushButton("📝  Note")
        self.btn_toggle_notes_code.setToolTip("Mostra/nascondi le note del codice")
        self.btn_toggle_notes_code.setCheckable(True)
        self.btn_toggle_notes_code.setChecked(False)
        self.btn_toggle_notes_code.clicked.connect(self._on_toggle_notes_code)
        code_info.addWidget(self.btn_toggle_notes_code)

        code_info.addStretch()
        hdr.addLayout(code_info, 1)
        layout.addLayout(hdr)

        # ---- Card PRT / ASM ----
        self.grp_prt_preview = QGroupBox("Parte / Assieme")
        prt_h = QHBoxLayout(self.grp_prt_preview)
        prt_h.setSpacing(8)

        self.lbl_prt_thumb = QLabel()
        self.lbl_prt_thumb.setFixedSize(QSize(_SUB_W, _SUB_H))
        self.lbl_prt_thumb.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_prt_thumb.setStyleSheet(
            "background-color: #181825; border: 1px solid #313244; border-radius: 6px;"
        )
        prt_h.addWidget(self.lbl_prt_thumb)

        prt_btns = QVBoxLayout()
        prt_btns.setSpacing(4)
        self.lbl_prt_info = QLabel()
        self.lbl_prt_info.setStyleSheet("font-size: 12px;")
        self.lbl_prt_info.setWordWrap(True)
        prt_btns.addWidget(self.lbl_prt_info)

        self.lbl_prt_checkout = QLabel()
        self.lbl_prt_checkout.setStyleSheet("font-size: 11px;")
        self.lbl_prt_checkout.setWordWrap(True)
        prt_btns.addWidget(self.lbl_prt_checkout)

        self.btn_prt_export_ws = QPushButton("📤  Esporta in WS")
        self.btn_prt_export_ws.setToolTip("Copia il file dall'archivio nella workspace locale")
        self.btn_prt_export_ws.clicked.connect(self._on_prt_export_ws)
        prt_btns.addWidget(self.btn_prt_export_ws)

        self.btn_prt_open_sw = QPushButton("🔨  Apri in SolidWorks")
        self.btn_prt_open_sw.clicked.connect(self._on_prt_open_sw)
        prt_btns.addWidget(self.btn_prt_open_sw)

        self.btn_create_sw = QPushButton("🔨  Crea in SW")
        self.btn_create_sw.setObjectName("btn_primary")
        self.btn_create_sw.clicked.connect(self._on_create_in_sw)
        prt_btns.addWidget(self.btn_create_sw)

        self.btn_create_file = QPushButton("📂  Crea da file")
        self.btn_create_file.clicked.connect(self._on_create_from_file)
        prt_btns.addWidget(self.btn_create_file)

        prt_btns.addStretch()
        prt_h.addLayout(prt_btns, 1)
        layout.addWidget(self.grp_prt_preview)

        # ---- Card DRW ----
        self.grp_drw_preview = QGroupBox("Disegno")
        drw_h = QHBoxLayout(self.grp_drw_preview)
        drw_h.setSpacing(8)

        self.lbl_drw_thumb = QLabel()
        self.lbl_drw_thumb.setFixedSize(QSize(_SUB_W, _SUB_H))
        self.lbl_drw_thumb.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_drw_thumb.setStyleSheet(
            "background-color: #181825; border: 1px solid #313244; border-radius: 6px;"
        )
        drw_h.addWidget(self.lbl_drw_thumb)

        drw_btns = QVBoxLayout()
        drw_btns.setSpacing(4)
        self.lbl_drw_info = QLabel()
        self.lbl_drw_info.setStyleSheet("font-size: 12px;")
        self.lbl_drw_info.setWordWrap(True)
        drw_btns.addWidget(self.lbl_drw_info)

        self.lbl_drw_checkout = QLabel()
        self.lbl_drw_checkout.setStyleSheet("font-size: 11px;")
        self.lbl_drw_checkout.setWordWrap(True)
        drw_btns.addWidget(self.lbl_drw_checkout)

        self.btn_drw_export_ws = QPushButton("📤  Esporta in WS")
        self.btn_drw_export_ws.setToolTip("Copia il disegno dall'archivio nella workspace locale")
        self.btn_drw_export_ws.clicked.connect(self._on_drw_export_ws)
        drw_btns.addWidget(self.btn_drw_export_ws)

        self.btn_drw_open_sw = QPushButton("🔨  Apri in SolidWorks")
        self.btn_drw_open_sw.clicked.connect(self._on_drw_open_sw)
        drw_btns.addWidget(self.btn_drw_open_sw)

        self.btn_add_drw = QPushButton("📐  Aggiungi DRW")
        self.btn_add_drw.clicked.connect(self._on_add_drw)
        drw_btns.addWidget(self.btn_add_drw)

        self.btn_create_drw_from_file = QPushButton("📂  Crea DRW da file")
        self.btn_create_drw_from_file.setToolTip("Importa un disegno esistente via SolidWorks")
        self.btn_create_drw_from_file.clicked.connect(self._on_create_drw_from_file)
        drw_btns.addWidget(self.btn_create_drw_from_file)

        drw_btns.addStretch()
        drw_h.addLayout(drw_btns, 1)
        layout.addWidget(self.grp_drw_preview)

        # ---- Workflow codice ----
        self.grp_code_workflow = QGroupBox("Workflow")
        wf_layout = QVBoxLayout(self.grp_code_workflow)
        wf_layout.setSpacing(6)
        wf_layout.setContentsMargins(8, 8, 8, 8)

        wf_state_row = QHBoxLayout()
        self.lbl_code_wf_state = QLabel()
        self.lbl_code_wf_state.setTextFormat(Qt.TextFormat.RichText)
        wf_state_row.addWidget(self.lbl_code_wf_state)
        self.lbl_code_wf_rev = QLabel()
        self.lbl_code_wf_rev.setStyleSheet("color: #a6adc8; font-size: 11px;")
        wf_state_row.addWidget(self.lbl_code_wf_rev)
        wf_state_row.addStretch()
        wf_layout.addLayout(wf_state_row)

        wf_btns_row = QHBoxLayout()
        self.btn_wf_release    = QPushButton("✅  Rilascia")
        self.btn_wf_release.setObjectName("btn_success")
        self.btn_wf_obsolete   = QPushButton("🗑  Rendi Obsoleto")
        self.btn_wf_obsolete.setObjectName("btn_warning")
        self.btn_wf_new_rev    = QPushButton("📋  Crea Revisione")
        self.btn_wf_cancel_rev = QPushButton("↩️  Annulla Revisione")
        self.btn_wf_release.clicked.connect(self._on_wf_release)
        self.btn_wf_obsolete.clicked.connect(self._on_wf_obsolete)
        self.btn_wf_new_rev.clicked.connect(self._on_wf_new_rev)
        self.btn_wf_cancel_rev.clicked.connect(self._on_wf_cancel_rev)
        for _b in (self.btn_wf_release, self.btn_wf_obsolete,
                   self.btn_wf_new_rev, self.btn_wf_cancel_rev):
            wf_btns_row.addWidget(_b)
            _b.setVisible(False)
        wf_btns_row.addStretch()
        wf_layout.addLayout(wf_btns_row)

        layout.addWidget(self.grp_code_workflow)

        # ---- Note codice ----
        self.grp_notes_code = QGroupBox("Note")
        grp_notes_layout = QVBoxLayout(self.grp_notes_code)
        grp_notes_layout.setContentsMargins(4, 4, 4, 4)
        grp_notes_layout.addWidget(self._build_notes_tab(doc_mode=False))
        self.grp_notes_code.setVisible(False)
        layout.addWidget(self.grp_notes_code)

        layout.addStretch()
        return w

    def load_code(self, code: str, docs: list[dict]):
        """Carica e mostra le informazioni del nodo codice (modalità codice)."""
        _same_code = (code == self._current_code)
        self._current_doc_id = None
        self._current_code = code
        self.stack.setCurrentIndex(1)

        # Header
        self.lbl_code_str.setText(code)
        rep = max(docs, key=lambda d: d["revision"]) if docs else None
        self.lbl_code_title.setText(rep.get("title", "") if rep else "")

        def _latest_active(dtype: str) -> "dict | None":
            """Restituisce l'ultima revisione attiva (non Obsoleto) del tipo dato,
            o l'ultima in assoluto se tutte obsolete."""
            pool = [d for d in docs if d["doc_type"] == dtype]
            if not pool:
                return None
            active = [d for d in pool if d["state"] != "Obsoleto"]
            return max(active or pool, key=lambda d: d["revision"])

        prt     = _latest_active("Parte")
        asm     = _latest_active("Assieme")
        drw     = _latest_active("Disegno")
        prt_asm = prt or asm

        # ---- Card PRT / ASM ----
        if prt_asm:
            self._code_prt_doc = prt_asm
            self._load_thumb_to(prt_asm, self.lbl_prt_thumb)
            icon = TYPE_ICON.get(prt_asm["doc_type"], "")
            rev  = prt_asm["revision"]
            has_archive = bool(prt_asm.get("archive_path"))
            in_ws = self._is_in_workspace(prt_asm)
            if has_archive:
                info       = f"{icon} Rev.{rev}  ✅  {prt_asm.get('state', '')}"
                info_style = "color: #a6e3a1; font-size: 12px;"
            elif prt_asm["is_locked"]:
                info       = f"{icon} Rev.{rev}  ⚠️  Modello non archiviato"
                info_style = "color: #f38ba8; font-size: 12px;"
            else:
                info       = f"{icon} Rev.{rev}  ⚠️  Modello non archiviato"
                info_style = "color: #f38ba8; font-size: 12px;"
            self.lbl_prt_info.setText(info)
            self.lbl_prt_info.setStyleSheet(info_style)
            if prt_asm["is_locked"]:
                locked_name = prt_asm.get("locked_by_name", "sconosciuto")
                locked_ws   = prt_asm.get("locked_ws", "")
                co_txt   = f"🔒 In checkout: {locked_name}"
                if locked_ws:
                    co_txt += f"  ({locked_ws})"
                co_style = "color: #fab387;"
            elif has_archive:
                co_txt   = "✅ Disponibile"
                co_style = "color: #a6e3a1;"
            else:
                co_txt   = ""
                co_style = ""
            self.lbl_prt_checkout.setText(co_txt)
            self.lbl_prt_checkout.setStyleSheet(f"font-size: 11px; {co_style}")
            self.btn_prt_export_ws.setVisible(has_archive and not in_ws)
            self.btn_prt_open_sw.setVisible(has_archive or in_ws)
            no_file = not has_archive and not prt_asm["is_locked"]
            self.btn_create_sw.setVisible(no_file)
            self.btn_create_file.setVisible(no_file)
        else:
            self._code_prt_doc = None
            self.lbl_prt_info.setText("⚠️  Modello non presente")
            self.lbl_prt_info.setStyleSheet("color: #585b70; font-size: 12px;")
            self.lbl_prt_checkout.setText("")
            self.btn_prt_export_ws.setVisible(False)
            self.btn_prt_open_sw.setVisible(False)
            self.btn_create_sw.setVisible(True)
            self.btn_create_file.setVisible(True)
        self.grp_prt_preview.setVisible(True)

        # ---- Card DRW ----
        if drw:
            self._code_drw_doc = drw
            self._load_thumb_to(drw, self.lbl_drw_thumb)
            rev = drw["revision"]
            has_archive = bool(drw.get("archive_path"))
            in_ws = self._is_in_workspace(drw)
            if has_archive:
                info       = f"📐 Rev.{rev}  ✅  {drw.get('state', '')}"
                info_style = "color: #a6e3a1; font-size: 12px;"
            else:
                info       = f"📐 Rev.{rev}  ⚠️  Disegno non archiviato"
                info_style = "color: #f38ba8; font-size: 12px;"
            self.lbl_drw_info.setText(info)
            self.lbl_drw_info.setStyleSheet(info_style)
            if drw["is_locked"]:
                locked_name = drw.get("locked_by_name", "sconosciuto")
                locked_ws   = drw.get("locked_ws", "")
                co_txt   = f"🔒 In checkout: {locked_name}"
                if locked_ws:
                    co_txt += f"  ({locked_ws})"
                co_style = "color: #fab387;"
            elif has_archive:
                co_txt   = "✅ Disponibile"
                co_style = "color: #a6e3a1;"
            else:
                co_txt   = ""
                co_style = ""
            self.lbl_drw_checkout.setText(co_txt)
            self.lbl_drw_checkout.setStyleSheet(f"font-size: 11px; {co_style}")
            self.btn_drw_export_ws.setVisible(has_archive and not in_ws)
            self.btn_drw_open_sw.setVisible(has_archive or in_ws)
            can_create_drw = not has_archive and not drw["is_locked"]
            self.btn_add_drw.setVisible(can_create_drw)
            self.btn_create_drw_from_file.setVisible(can_create_drw)
        else:
            self._code_drw_doc = None
            self.lbl_drw_thumb.setPixmap(QPixmap())
            self.lbl_drw_thumb.setText("📐")
            self.lbl_drw_thumb.setStyleSheet(
                "background-color: #181825; border: 1px solid #313244; "
                "border-radius: 6px; font-size: 40px; color: #585b70;"
            )
            self.lbl_drw_info.setText("⚠️  Disegno non presente")
            self.lbl_drw_info.setStyleSheet("color: #585b70; font-size: 12px;")
            self.lbl_drw_checkout.setText("")
            self.btn_drw_export_ws.setVisible(False)
            self.btn_drw_open_sw.setVisible(False)
            self.btn_add_drw.setVisible(True)
            self.btn_create_drw_from_file.setVisible(True)
        self.grp_drw_preview.setVisible(True)

        # ---- Riferimenti per handlers ----
        prt_asm_no_file = [
            d for d in docs
            if d["doc_type"] in ("Parte", "Assieme")
            and not d.get("archive_path") and not d["is_locked"]
        ]
        prt_asm_with_file = [
            d for d in docs
            if d["doc_type"] in ("Parte", "Assieme") and d.get("archive_path")
        ]
        self._code_action_doc_id   = prt_asm_no_file[0]["id"]   if prt_asm_no_file   else None
        self._code_prt_asm_for_drw = prt_asm_with_file[0]["id"] if prt_asm_with_file else None

        # ---- Workflow sezione codice ----
        self._code_docs = docs
        self._code_representative = self._get_code_representative_doc(docs)
        self._update_code_workflow_section()

        # ---- Note ----
        if not _same_code:
            self._load_notes_to(self._notes_edit_code, self._notes_lbl_updated_code, code)
        self._notes_edit_code.setReadOnly(False)

    def _on_create_in_sw(self):
        if self._code_action_doc_id:
            self.create_in_sw_requested.emit(self._code_action_doc_id)

    def _on_create_from_file(self):
        if self._code_action_doc_id:
            self.create_from_file_requested.emit(self._code_action_doc_id)

    def _on_add_drw(self):
        if self._code_prt_asm_for_drw:
            self.add_drw_requested.emit(self._code_prt_asm_for_drw)

    def _on_create_drw_from_file(self):
        if self._code_prt_asm_for_drw:
            self.create_drw_from_file_requested.emit(self._code_prt_asm_for_drw)

    def _on_doc_add_drw(self):
        if self._current_doc_id:
            self.add_drw_requested.emit(self._current_doc_id)

    def _on_doc_create_drw_from_file(self):
        if self._current_doc_id:
            self.create_drw_from_file_requested.emit(self._current_doc_id)

    # ------------------------------------------------------------------
    #  Workflow codice
    # ------------------------------------------------------------------
    def _update_code_workflow_section(self):
        """Aggiorna badge stato e bottoni workflow nella sezione codice."""
        rep = self._code_representative
        if rep:
            state = rep["state"]
            badge = STATE_BADGE_STYLE.get(state, "")
            self.lbl_code_wf_state.setText(
                f"<span style='{badge}'>&nbsp;{state}&nbsp;</span>"
            )
            self.lbl_code_wf_rev.setText(
                f"  Rev.{rep['revision']}  ({rep['doc_type']})"
            )
            available  = session.workflow.get_available_transitions(state)
            is_latest  = session.workflow.is_latest_revision(rep)
            is_prt_asm = rep["doc_type"] in ("Parte", "Assieme")

            can_release  = "Rilasciato" in available and session.can("release")
            can_obsolete = "Obsoleto"   in available and session.can("release")
            can_new_rev  = (state == "Rilasciato" and is_latest
                            and session.can("create"))
            can_cancel   = (state == "In Revisione" and not rep["is_locked"]
                            and session.can("create"))

            self.btn_wf_release.setVisible(can_release)
            self.btn_wf_obsolete.setVisible(can_obsolete)
            self.btn_wf_new_rev.setVisible(can_new_rev)
            self.btn_wf_cancel_rev.setVisible(can_cancel)
        else:
            self.lbl_code_wf_state.setText(
                "<span style='color:#585b70;'>Nessun documento attivo</span>"
            )
            self.lbl_code_wf_rev.setText("")
            for _b in (self.btn_wf_release, self.btn_wf_obsolete,
                       self.btn_wf_new_rev, self.btn_wf_cancel_rev):
                _b.setVisible(False)

    def _reload_code_panel(self):
        """Ricarica il pannello codice corrente senza full tree refresh."""
        if not self._current_code or not session.is_connected:
            return
        docs = [d for d in session.files.search_documents(code=self._current_code)
                if d["code"] == self._current_code]
        self.load_code(self._current_code, docs)

    def _on_wf_release(self):
        rep = self._code_representative
        if rep:
            self.workflow_requested.emit(rep, list(self._code_docs), "Rilasciato")

    def _on_wf_obsolete(self):
        rep = self._code_representative
        if rep:
            self.workflow_requested.emit(rep, list(self._code_docs), "Obsoleto")

    def _on_wf_new_rev(self):
        rep = self._code_representative
        if not rep:
            return
        doc = rep

        try:
            current_num = int(doc["revision"])
            next_rev = str(current_num + 1).zfill(len(doc["revision"]))
        except ValueError:
            from PyQt6.QtWidgets import QInputDialog
            next_rev, ok = QInputDialog.getText(
                self, "Nuova revisione",
                f"Revisione attuale: {doc['revision']}\nInserisci nuova revisione:"
            )
            if not ok or not next_rev.strip():
                return
            next_rev = next_rev.strip()

        existing = session.db.fetchone(
            "SELECT id FROM documents WHERE code=? AND revision=?",
            (doc["code"], next_rev),
        )
        if existing:
            QMessageBox.warning(
                self, "Revisione esistente",
                f"La revisione {next_rev} esiste già per il codice {doc['code']}."
            )
            return

        r = QMessageBox.question(
            self, "Nuova revisione",
            f"Creare revisione {next_rev} dal codice {doc['code']} rev.{doc['revision']}?\n\n"
            "Il file archiviato verrà copiato come base per la nuova revisione.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if r != QMessageBox.StandardButton.Yes:
            return

        try:
            new_id = session.workflow.new_revision(
                doc["id"], session.user["id"], next_rev,
                shared_paths=session.sp,
            )
            QMessageBox.information(
                self, "Nuova revisione",
                f"Creata revisione {next_rev} (ID: {new_id})"
            )
            self._reload_code_panel()
            self.workflow_action_completed.emit()
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

    def _on_wf_cancel_rev(self):
        rep = self._code_representative
        if not rep:
            return
        doc = rep
        r = QMessageBox.question(
            self, "Annulla revisione",
            f"Annullare la revisione {doc['revision']} del codice {doc['code']}?\n\n"
            "Il documento e il relativo file archiviato verranno eliminati.\n"
            "La revisione precedente resterà inalterata.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if r != QMessageBox.StandardButton.Yes:
            return
        try:
            session.workflow.cancel_revision(
                doc["id"], session.user["id"],
                shared_paths=session.sp,
            )
            QMessageBox.information(
                self, "Annulla revisione",
                f"Revisione {doc['revision']} annullata."
            )
            self._reload_code_panel()
            self.workflow_action_completed.emit()
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

    def _on_prt_export_ws(self):
        self._export_doc_to_ws(self._code_prt_doc)

    def _on_drw_export_ws(self):
        self._export_doc_to_ws(self._code_drw_doc)

    def _on_prt_open_sw(self):
        self._open_in_solidworks(self._code_prt_doc)

    def _on_drw_open_sw(self):
        self._open_in_solidworks(self._code_drw_doc)

    # ------------------------------------------------------------------
    #  Helper codice – archivio / workspace
    # ------------------------------------------------------------------
    def _get_archive_file(self, doc: dict) -> "Path | None":
        """Restituisce il path completo del file in archivio, o None se assente."""
        if not session.sp or not doc.get("archive_path"):
            return None
        p = session.sp.root / doc["archive_path"]
        return p if p.exists() else None

    def _is_in_workspace(self, doc: dict) -> bool:
        """Verifica se il file esiste già nella workspace SolidWorks configurata."""
        from config import load_local_config
        ws = load_local_config().get("sw_workspace", "")
        if not ws:
            return False
        ext = _EXT_FOR_TYPE.get(doc.get("doc_type", ""), "")
        return (Path(ws) / (doc.get("code", "") + ext)).exists()

    def _export_doc_to_ws(self, doc: "dict | None"):
        """Copia il file dall'archivio nella workspace locale."""
        if not doc:
            return
        archive_file = self._get_archive_file(doc)
        if not archive_file:
            QMessageBox.warning(self, "File non disponibile",
                                "Il file non è presente nell'archivio.")
            return
        from config import load_local_config
        ws = load_local_config().get("sw_workspace", "")
        if not ws:
            QMessageBox.warning(
                self, "Workspace non configurata",
                "Configurare la workspace in Strumenti → Configurazione SolidWorks."
            )
            return
        ws_path = Path(ws)
        ws_path.mkdir(parents=True, exist_ok=True)
        dest = ws_path / archive_file.name
        if dest.exists():
            QMessageBox.information(self, "Già presente",
                                    f"Il file è già nella workspace:\n{dest}")
            return
        try:
            shutil.copy2(archive_file, dest)
            QMessageBox.information(self, "Esportazione completata",
                                    f"File copiato in:\n{dest}")
            # Ricarica il pannello per aggiornare visibilità pulsanti
            if session.files and self._current_code:
                docs = [d for d in session.files.search_documents(code=self._current_code)
                        if d["code"] == self._current_code]
                self.load_code(self._current_code, docs)
        except Exception as e:
            QMessageBox.critical(self, "Errore esportazione", str(e))

    def _open_in_solidworks(self, doc: "dict | None"):
        """Apre il file in SolidWorks (cerca prima in WS, poi in archivio)."""
        if not doc:
            return
        from config import load_local_config
        ws = load_local_config().get("sw_workspace", "")
        ext = _EXT_FOR_TYPE.get(doc.get("doc_type", ""), "")
        target: "Path | None" = None
        if ws:
            ws_file = Path(ws) / (doc.get("code", "") + ext)
            if ws_file.exists():
                target = ws_file
        if not target:
            target = self._get_archive_file(doc)
        if not target:
            QMessageBox.warning(
                self, "File non trovato",
                "Il file non è disponibile nell'archivio né nella workspace."
            )
            return

        doc_type_map = {"Parte": 1, "Assieme": 2, "Disegno": 3}
        type_id = doc_type_map.get(doc.get("doc_type", "Parte"), 1)
        self._open_in_solidworks_path(target, doc.get("doc_type", "Parte"))

    def _open_in_solidworks_path(self, target: "Path", doc_type: str):
        """Apre un file in SolidWorks dato il path completo e il tipo documento."""
        doc_type_map = {"Parte": 1, "Assieme": 2, "Disegno": 3}
        type_id = doc_type_map.get(doc_type, 1)
        try:
            import win32com.client
            try:
                sw = win32com.client.GetActiveObject("SldWorks.Application")
            except Exception:
                sw = win32com.client.Dispatch("SldWorks.Application")
            sw.Visible = True
            sw.OpenDoc(str(target).replace("/", "\\"), type_id)
            return
        except Exception as e:
            try:
                from ui.sw_config_dialog import SWConfigDialog
                sw_exe = SWConfigDialog.get_solidworks_exe()
                if sw_exe and Path(str(sw_exe)).exists():
                    subprocess.Popen([str(sw_exe), str(target)])
                else:
                    subprocess.Popen(["cmd", "/c", "start", "", str(target)])
            except Exception:
                QMessageBox.critical(self, "Errore apertura SolidWorks", str(e))

    # ------------------------------------------------------------------
    #  Revisioni obsolete — apertura file
    # ------------------------------------------------------------------
    def _get_selected_revision_doc(self) -> "dict | None":
        """Restituisce il documento della riga selezionata nel tab Revisioni."""
        if not self.tbl_revisions.selectedItems():
            return None
        row = self.tbl_revisions.currentRow()
        item = self.tbl_revisions.item(row, 0)
        if not item:
            return None
        doc_id = item.data(Qt.ItemDataRole.UserRole)
        return session.files.get_document(doc_id) if doc_id and session.files else None

    def _get_pdf_path_for_doc(self, doc: dict) -> "Path | None":
        """Restituisce il Path del PDF di un documento specifico, o None."""
        pdf_path = doc.get("pdf_path") or ""
        if not pdf_path:
            return None
        if session.sp:
            p = session.sp.root / pdf_path
            if p.exists():
                return p
        p_abs = Path(pdf_path)
        return p_abs if p_abs.exists() else None

    def _on_revision_selection_changed(self):
        """Abilita/disabilita i pulsanti del tab Revisioni in base alla riga selezionata."""
        doc = self._get_selected_revision_doc()
        has_archive = bool(
            doc and doc.get("archive_path") and session.sp
            and (session.sp.root / doc["archive_path"]).exists()
        )
        has_pdf = bool(doc and self._get_pdf_path_for_doc(doc))
        self.btn_rev_edrawings.setEnabled(has_archive)
        self.btn_rev_sw.setEnabled(has_archive)
        self.btn_rev_pdf.setEnabled(has_pdf)

    def _open_obsolete(self, action: str):
        """Copia il file della revisione obsoleta nella cartella OBSOLETI e lo apre."""
        doc = self._get_selected_revision_doc()
        if not doc:
            return

        from config import load_local_config
        ws = load_local_config().get("sw_workspace", "")
        if not ws:
            QMessageBox.warning(
                self, "Workspace non configurata",
                "Configurare la workspace locale in Strumenti → Impostazioni\n"
                "prima di aprire i file."
            )
            return

        obsoleti_dir = Path(ws) / "OBSOLETI"
        try:
            obsoleti_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            QMessageBox.critical(self, "Errore cartella OBSOLETI", str(e))
            return

        if action in ("edrawings", "sw"):
            src = session.sp.root / doc["archive_path"]
            if not src.exists():
                QMessageBox.warning(
                    self, "File non trovato",
                    f"Il file non è stato trovato nell'archivio:\n{src}"
                )
                return
            dest = obsoleti_dir / src.name
            try:
                shutil.copy2(src, dest)
            except Exception as e:
                QMessageBox.critical(self, "Errore copia file", str(e))
                return

            if action == "edrawings":
                from ui.sw_config_dialog import SWConfigDialog
                exe = SWConfigDialog.get_edrawings_exe()
                if exe:
                    subprocess.Popen([str(exe), str(dest)])
                else:
                    QMessageBox.warning(
                        self, "eDrawings non configurato",
                        "Eseguibile eDrawings non configurato.\n\n"
                        "Aprire Strumenti → Configurazione SolidWorks → tab SolidWorks\n"
                        "e impostare il percorso di eDrawings."
                    )
            else:
                self._open_in_solidworks_path(dest, doc.get("doc_type", "Parte"))

        elif action == "pdf":
            src = self._get_pdf_path_for_doc(doc)
            if not src:
                QMessageBox.warning(self, "PDF non trovato", "Il file PDF non è disponibile.")
                return
            dest = obsoleti_dir / src.name
            try:
                shutil.copy2(src, dest)
            except Exception as e:
                QMessageBox.critical(self, "Errore copia PDF", str(e))
                return
            import os
            os.startfile(str(dest))

    # ------------------------------------------------------------------
    #  Thumbnail
    # ------------------------------------------------------------------
    def _load_thumb_to(self, doc: dict, lbl: QLabel):
        """Carica thumbnail in un QLabel specifico (con fallback icona tipo)."""
        thumb_path = self._get_thumbnail_path(doc)
        _w = lbl.width()  or _SUB_W
        _h = lbl.height() or _SUB_H
        if thumb_path and thumb_path.exists():
            pixmap = QPixmap(str(thumb_path))
            if not pixmap.isNull():
                scaled = pixmap.scaled(
                    _w, _h,
                    Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation,
                )
                lbl.setPixmap(scaled)
                lbl.setStyleSheet(
                    "background-color: #181825; border: 1px solid #313244;"
                    " border-radius: 6px;"
                )
                return
        # Fallback: icona tipo
        icon = _FALLBACK_ICONS.get(doc["doc_type"], "📄")
        lbl.setPixmap(QPixmap())
        lbl.setText(icon)
        lbl.setStyleSheet(
            "background-color: #181825; border: 1px solid #313244; "
            "border-radius: 6px; font-size: 36px; color: #585b70;"
        )

    def _load_thumbnail(self, doc: dict):
        """Carica thumbnail nella label principale (modalità documento)."""
        thumb_path = self._get_thumbnail_path(doc)

        if thumb_path and thumb_path.exists():
            pixmap = QPixmap(str(thumb_path))
            if not pixmap.isNull():
                scaled = pixmap.scaled(
                    _THUMB_W, _THUMB_H,
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
        self.lbl_thumb.setPixmap(QPixmap())
        self.lbl_thumb.setText(icon)
        self.lbl_thumb.setStyleSheet(
            "background-color: #181825; border: 1px solid #313244; "
            "border-radius: 6px; font-size: 64px; color: #585b70;"
        )

    def _get_thumbnail_path(self, doc: dict) -> "Path | None":
        """Restituisce il path della thumbnail solo per documenti codificati (file archiviato)."""
        if not session.sp:
            return None
        if not doc.get("archive_path"):
            return None
        code = doc.get("code", "")
        rev = doc.get("revision", "")
        if not code:
            return None
        suffix = "_DRW" if doc.get("doc_type") == "Disegno" else ""
        thumb_file = session.sp.thumbnails / f"{code}_{rev}{suffix}.png"
        return thumb_file

    # ------------------------------------------------------------------
    #  Sync proprietà SW ↔ PDM
    # ------------------------------------------------------------------
    def _get_sw_file_for_sync(self) -> "Path | None":
        """
        Restituisce il path del file SW del documento corrente.
        Cerca prima nella workspace, poi nell'archivio.
        Funziona per PRT, ASM e DRW (ognuno cerca il proprio file).
        """
        doc_id = self._current_doc_id
        if not doc_id or not session.is_connected:
            return None

        doc = session.files.get_document(doc_id)
        if not doc:
            return None

        from config import load_local_config
        from core.file_manager import EXT_FOR_TYPE
        ws = load_local_config().get("sw_workspace", "")
        ext = EXT_FOR_TYPE.get(doc.get("doc_type", ""), ".SLDPRT")

        # 1) Workspace
        if ws:
            ws_file = Path(ws) / (doc["code"] + ext)
            if ws_file.exists():
                return ws_file

        # 2) Archivio
        if session.sp and doc.get("archive_path"):
            arch = session.sp.root / doc["archive_path"]
            if arch.exists():
                return arch

        return None

    def _refresh_props_table(self):
        """Ricarica la tabella proprietà e, se in modalità documento, ricarica il pannello."""
        if not self._current_doc_id:
            return
        self.tbl_props.setRowCount(0)
        props = session.properties.get_properties(self._current_doc_id)
        for name, value in props.items():
            row = self.tbl_props.rowCount()
            self.tbl_props.insertRow(row)
            self.tbl_props.setItem(row, 0, QTableWidgetItem(name))
            self.tbl_props.setItem(row, 1, QTableWidgetItem(value))

    def _on_sync_props(self):
        """Aggiorna Proprietà: sincronizza PDM ↔ SW secondo la mappatura."""
        doc_id = self._current_doc_id
        if not doc_id:
            return
        file_path = self._get_sw_file_for_sync()
        if not file_path:
            QMessageBox.warning(
                self, "File non trovato",
                "Nessun file SW disponibile per la sincronizzazione.\n"
                "Aprire il file in SolidWorks o esportarlo nella workspace."
            )
            return
        try:
            result = session.properties.sync_bidirectional(doc_id, file_path)
            if not result.get("ok") and result.get("error"):
                QMessageBox.warning(self, "Sincronizzazione", result["error"])
                return
            self._refresh_props_table()
            msg = (
                f"Sincronizzazione completata.\n"
                f"Proprietà importate da SW: {result.get('imported_count', 0)}\n"
                f"Proprietà scritte in SW: {result.get('written_count', 0)}"
            )
            if result.get("updated_owner"):
                msg += "\nTitolo/descrizione aggiornati nel PDM."
            QMessageBox.information(self, "Aggiorna Proprietà", msg)
        except Exception as e:
            QMessageBox.critical(self, "Errore sincronizzazione", str(e))

    def _on_import_from_sw(self):
        """Importa da SW: legge le proprietà dal file SW e le salva nel PDM."""
        doc_id = self._current_doc_id
        if not doc_id:
            return
        file_path = self._get_sw_file_for_sync()
        if not file_path:
            QMessageBox.warning(
                self, "File non trovato",
                "Nessun file SW disponibile per l'importazione.\n"
                "Aprire il file in SolidWorks o esportarlo nella workspace."
            )
            return
        try:
            result = session.properties.sync_sw_to_pdm(doc_id, file_path)
            if not result.get("ok"):
                QMessageBox.warning(self, "Importazione", result.get("error", "Errore sconosciuto"))
                return
            self._refresh_props_table()
            msg = (
                f"Importazione completata.\n"
                f"Proprietà importate: {result.get('imported_count', 0)}"
            )
            if result.get("updated_owner"):
                msg += "\nTitolo/descrizione aggiornati nel PDM."
            QMessageBox.information(self, "Importa da SW", msg)
        except Exception as e:
            QMessageBox.critical(self, "Errore importazione", str(e))

    def _on_export_to_sw(self):
        """Esporta in SW: scrive i campi PDM nel file SW secondo la mappatura.
        Per i DRW scrive i campi del documento padre (PRT/ASM) nel file DRW."""
        doc_id = self._current_doc_id
        if not doc_id:
            return

        file_path = self._get_sw_file_for_sync()
        if not file_path:
            QMessageBox.warning(
                self, "File non trovato",
                "Nessun file SW disponibile per l'esportazione.\n"
                "Il file deve essere aperto in SolidWorks o presente nella workspace."
            )
            return
        try:
            result = session.properties.sync_pdm_to_sw(doc_id, file_path)
            if not result.get("ok"):
                QMessageBox.warning(self, "Esportazione", result.get("error", "Errore sconosciuto"))
                return
            msg = f"Esportazione completata.\nProprietà scritte in SW: {result.get('written_count', 0)}"
            QMessageBox.information(self, "Esporta in SW", msg)
        except Exception as e:
            QMessageBox.critical(self, "Errore esportazione", str(e))

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

    # ------------------------------------------------------------------
    #  Toggle Note tab
    # ------------------------------------------------------------------
    def _on_toggle_notes(self, checked: bool):
        self.tabs.setTabVisible(self._notes_tab_index, checked)
        if checked:
            self.tabs.setCurrentIndex(self._notes_tab_index)

    def _on_toggle_notes_code(self, checked: bool):
        self.grp_notes_code.setVisible(checked)

    # ------------------------------------------------------------------
    #  Note per codice
    # ------------------------------------------------------------------
    def _load_notes_to(self, edit: QTextEdit, lbl: QLabel, code: str):
        """Carica la nota del codice nell'editor specificato."""
        if not session.is_connected or not code:
            edit.setPlainText("")
            lbl.setText("")
            return
        try:
            row = session.db.get_note(code)
        except Exception:
            row = None
        if row and row.get("content"):
            edit.setHtml(row["content"])
        else:
            edit.setPlainText("")
        if row:
            by = row.get("updated_by_name") or ""
            at = (row.get("updated_at") or "")[:16].replace("T", " ")
            lbl.setText(f"Ultimo aggiorn.: {at}  {by}".strip())
        else:
            lbl.setText("")

    def _save_notes_doc(self):
        self._save_notes_from(self._notes_edit, self._notes_lbl_updated)

    def _save_notes_code(self):
        self._save_notes_from(self._notes_edit_code, self._notes_lbl_updated_code)

    def _save_notes_from(self, edit: QTextEdit, lbl: QLabel):
        """Salva la nota dal QTextEdit specificato."""
        code = self._current_code
        if not session.is_connected or not code:
            return
        content = edit.toHtml()
        user_id = (session.user or {}).get("id", 0)
        try:
            session.db.save_note(code, content, user_id)
            row = session.db.get_note(code)
            if row:
                by = row.get("updated_by_name") or ""
                at = (row.get("updated_at") or "")[:16].replace("T", " ")
                lbl.setText(f"Ultimo aggiorn.: {at}  {by}".strip())
        except Exception as e:
            QMessageBox.critical(self, "Errore salvataggio nota", str(e))

    def _insert_date_user(self, edit: QTextEdit):
        """Inserisce dd/mm/yyyy - Nome Utente: nel cursore corrente."""
        today = date.today().strftime("%d/%m/%Y")
        user_name = (session.user or {}).get("full_name", "")
        text = f"{today} - {user_name}: " if user_name else f"{today}: "
        edit.insertPlainText(text)

    # ------------------------------------------------------------------
    #  PDF
    # ------------------------------------------------------------------
    def _current_pdf_path(self) -> "Path | None":
        """Restituisce il Path del PDF del documento corrente, o None."""
        if not self._current_doc_id or not session.is_connected:
            return None
        doc = session.files.get_document(self._current_doc_id)
        if not doc:
            return None
        pdf_path = doc.get("pdf_path") or ""
        if not pdf_path:
            return None
        # Prova come percorso relativo (nuovo formato)
        if session.sp:
            p = session.sp.root / pdf_path
            if p.exists():
                return p
        # Fallback percorso assoluto (vecchio formato)
        p_abs = Path(pdf_path)
        return p_abs if p_abs.exists() else None

    def _on_open_pdf(self):
        """Copia il PDF nella workspace e lo apre con il visualizzatore predefinito."""
        p = self._current_pdf_path()
        if not p:
            QMessageBox.warning(self, "PDF non trovato", "Il file PDF non è disponibile.")
            return
        import os
        from config import load_local_config
        ws = load_local_config().get("sw_workspace", "")
        dest = p  # fallback: apri direttamente dall'archivio
        if ws:
            ws_path = Path(ws)
            ws_path.mkdir(parents=True, exist_ok=True)
            ws_dest = ws_path / p.name
            try:
                shutil.copy2(p, ws_dest)
                dest = ws_dest
            except Exception:
                pass  # fallback: apri dall'archivio
        os.startfile(str(dest))

    def _on_save_pdf(self):
        """Copia il PDF in una cartella scelta dall'utente."""
        p = self._current_pdf_path()
        if not p:
            QMessageBox.warning(self, "PDF non trovato", "Il file PDF non è disponibile.")
            return
        from PyQt6.QtWidgets import QFileDialog
        dest_str, _ = QFileDialog.getSaveFileName(
            self,
            "Salva copia PDF",
            p.name,
            "PDF (*.pdf)",
        )
        if not dest_str:
            return
        try:
            shutil.copy2(p, dest_str)
            QMessageBox.information(self, "PDF salvato", f"Copia salvata in:\n{dest_str}")
        except Exception as e:
            QMessageBox.critical(self, "Errore salvataggio", str(e))

    # ------------------------------------------------------------------
    def _on_generate_pdf(self):
        """Genera PDF del disegno corrente in background."""
        if not self._current_doc_id or not session.is_connected:
            return
        doc = session.files.get_document(self._current_doc_id)
        if not doc or not doc.get("archive_path"):
            return
        if not session.sp:
            return

        src  = session.sp.root / doc["archive_path"]
        dest = src.with_suffix(".pdf")

        self.btn_gen_pdf.setEnabled(False)
        self.lbl_pdf_path.setText("Generazione in corso…")

        import threading
        threading.Thread(
            target=self._generate_pdf_worker,
            args=(self._current_doc_id, src, dest),
            daemon=True,
        ).start()

    def _generate_pdf_worker(self, doc_id: int, src: Path, dest: Path):
        """Worker thread per la generazione PDF (subprocess pdf_worker.py)."""
        import sys
        import tempfile
        import os
        _WORKER = Path(__file__).parent.parent / "core" / "pdf_worker.py"
        _TIMEOUT = 120

        log_text = ""
        try:
            tmp_fd, tmp_log = tempfile.mkstemp(suffix=".txt", prefix="pdf_")
            os.close(tmp_fd)
            with open(tmp_log, "w") as lf:
                proc = subprocess.Popen(
                    [sys.executable, str(_WORKER), str(src), str(dest)],
                    stdout=lf,
                    stderr=lf,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
            proc.wait(timeout=_TIMEOUT)
            log_text = Path(tmp_log).read_text(errors="replace").strip()
            ok = proc.returncode == 0 and dest.exists()
        except Exception as e:
            log_text = str(e)
            ok = False

        if ok and session.db and session.sp:
            try:
                rel_path = str(dest.relative_to(session.sp.root))
                session.db.set_pdf_path(doc_id, rel_path)
            except Exception:
                pass

        # Aggiorna UI in modo thread-safe tramite segnale Qt
        detail = str(dest) if ok else log_text
        self._pdf_done.emit(ok, detail)

    def _on_pdf_done(self, ok: bool, pdf_path: str):
        """Chiamato dal thread principale dopo generazione PDF.
        pdf_path = percorso PDF se ok=True, altrimenti log diagnostico.
        """
        self.btn_gen_pdf.setEnabled(True)
        if ok:
            self.lbl_pdf_path.setText(pdf_path)
            self.btn_open_pdf.setEnabled(True)
            self.btn_save_pdf.setEnabled(True)
            QMessageBox.information(self, "PDF generato", f"PDF salvato in:\n{pdf_path}")
        else:
            self.lbl_pdf_path.setText("Errore nella generazione PDF")
            self.btn_open_pdf.setEnabled(False)
            self.btn_save_pdf.setEnabled(False)
            # Mostra il log del subprocess per diagnostica
            detail = pdf_path[:600] if pdf_path else "(nessun dettaglio)"
            QMessageBox.warning(
                self, "Errore PDF",
                "Impossibile generare il PDF.\n\n"
                f"Dettaglio:\n{detail}"
            )

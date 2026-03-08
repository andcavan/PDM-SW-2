# =============================================================================
#  ui/detail_panel.py  –  Pannello dettaglio documento (readonly, anteprima)
# =============================================================================
import shutil
import subprocess
from pathlib import Path

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTabWidget, QTableWidget, QTableWidgetItem, QHeaderView,
    QGroupBox, QFormLayout, QMessageBox, QSizePolicy,
    QStackedWidget
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

    def __init__(self, parent=None):
        super().__init__(parent)
        self._current_doc_id: int | None = None
        self._current_code:   str | None = None
        self._code_action_doc_id:    int | None = None  # PRT/ASM senza file
        self._code_prt_asm_for_drw:  int | None = None  # PRT/ASM con file (per DRW)
        self._code_prt_doc:  dict | None = None
        self._code_drw_doc:  dict | None = None
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
        self.grp_doc_drw.setVisible(False)
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

        # ---- Card PRT / ASM ----
        self.grp_prt_preview = QGroupBox("Parte / Assieme")
        prt_h = QHBoxLayout(self.grp_prt_preview)
        prt_h.setSpacing(8)

        self.lbl_prt_thumb = QLabel()
        self.lbl_prt_thumb.setFixedSize(QSize(110, 110))
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

        self.btn_prt_export_ws = QPushButton("📤  Esporta in WS")
        self.btn_prt_export_ws.setToolTip("Copia il file dall'archivio nella workspace locale")
        self.btn_prt_export_ws.clicked.connect(self._on_prt_export_ws)
        prt_btns.addWidget(self.btn_prt_export_ws)

        self.btn_prt_open_sw = QPushButton("🔨  Apri in SolidWorks")
        self.btn_prt_open_sw.clicked.connect(self._on_prt_open_sw)
        prt_btns.addWidget(self.btn_prt_open_sw)

        prt_btns.addStretch()
        prt_h.addLayout(prt_btns, 1)
        layout.addWidget(self.grp_prt_preview)

        # ---- Card DRW ----
        self.grp_drw_preview = QGroupBox("Disegno")
        drw_h = QHBoxLayout(self.grp_drw_preview)
        drw_h.setSpacing(8)

        self.lbl_drw_thumb = QLabel()
        self.lbl_drw_thumb.setFixedSize(QSize(110, 110))
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

        # ---- Azioni creazione (quando non c'è ancora il file PRT/ASM) ----
        self.grp_actions = QGroupBox("Azioni")
        act_layout = QVBoxLayout(self.grp_actions)

        self.btn_create_sw = QPushButton("🔨  Crea in SW")
        self.btn_create_sw.setObjectName("btn_primary")
        self.btn_create_sw.clicked.connect(self._on_create_in_sw)
        act_layout.addWidget(self.btn_create_sw)

        self.btn_create_file = QPushButton("📂  Crea da file")
        self.btn_create_file.clicked.connect(self._on_create_from_file)
        act_layout.addWidget(self.btn_create_file)

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

        prt = next((d for d in docs if d["doc_type"] == "Parte"),   None)
        asm = next((d for d in docs if d["doc_type"] == "Assieme"), None)
        drw = next((d for d in docs if d["doc_type"] == "Disegno"), None)
        prt_asm = prt or asm

        # ---- Card PRT / ASM ----
        if prt_asm:
            self._code_prt_doc = prt_asm
            self._load_thumb_to(prt_asm, self.lbl_prt_thumb)
            icon = TYPE_ICON.get(prt_asm["doc_type"], "")
            rev  = prt_asm["revision"]
            if prt_asm.get("archive_path"):
                info       = f"{icon} Rev.{rev}  ✅  {prt_asm.get('state', '')}"
                info_style = "color: #a6e3a1; font-size: 12px;"
            elif prt_asm["is_locked"]:
                info       = f"{icon} Rev.{rev}  🔒 In checkout"
                info_style = "color: #fab387; font-size: 12px;"
            else:
                info       = f"{icon} Rev.{rev}  ❌ Non archiviato"
                info_style = "color: #f38ba8; font-size: 12px;"
            self.lbl_prt_info.setText(info)
            self.lbl_prt_info.setStyleSheet(info_style)
            has_archive = bool(prt_asm.get("archive_path"))
            in_ws = self._is_in_workspace(prt_asm)
            self.btn_prt_export_ws.setVisible(has_archive and not in_ws)
            self.btn_prt_open_sw.setVisible(has_archive or in_ws)
            self.grp_prt_preview.setVisible(True)
        else:
            self._code_prt_doc = None
            self.grp_prt_preview.setVisible(False)

        # ---- Card DRW ----
        prt_asm_archived = bool(prt_asm and prt_asm.get("archive_path"))
        if drw:
            self._code_drw_doc = drw
            self._load_thumb_to(drw, self.lbl_drw_thumb)
            rev = drw["revision"]
            if drw.get("archive_path"):
                info       = f"📐 Rev.{rev}  ✅  {drw.get('state', '')}"
                info_style = "color: #a6e3a1; font-size: 12px;"
            elif drw["is_locked"]:
                info       = f"📐 Rev.{rev}  🔒 In checkout"
                info_style = "color: #fab387; font-size: 12px;"
            else:
                info       = f"📐 Rev.{rev}  ❌ Non archiviato"
                info_style = "color: #f38ba8; font-size: 12px;"
            self.lbl_drw_info.setText(info)
            self.lbl_drw_info.setStyleSheet(info_style)
            has_archive = bool(drw.get("archive_path"))
            in_ws = self._is_in_workspace(drw)
            self.btn_drw_export_ws.setVisible(has_archive and not in_ws)
            self.btn_drw_open_sw.setVisible(has_archive or in_ws)
            # Mostra "Aggiungi DRW" e "Crea DRW da file" quando non c'è ancora il file
            can_create_drw = not has_archive and not drw["is_locked"]
            self.btn_add_drw.setVisible(can_create_drw)
            self.btn_create_drw_from_file.setVisible(can_create_drw)
            self.grp_drw_preview.setVisible(True)
        else:
            self._code_drw_doc = None
            self.lbl_drw_thumb.setPixmap(QPixmap())
            self.lbl_drw_thumb.setText("📐")
            self.lbl_drw_thumb.setStyleSheet(
                "background-color: #181825; border: 1px solid #313244; "
                "border-radius: 6px; font-size: 40px; color: #585b70;"
            )
            self.lbl_drw_info.setText(
                "Disegno non ancora creato" if prt_asm_archived else "Nessun disegno"
            )
            self.lbl_drw_info.setStyleSheet("color: #585b70; font-size: 12px;")
            self.btn_drw_export_ws.setVisible(False)
            self.btn_drw_open_sw.setVisible(False)
            self.btn_add_drw.setVisible(prt_asm_archived)
            self.btn_create_drw_from_file.setVisible(prt_asm_archived)
            self.grp_drw_preview.setVisible(prt_asm_archived)

        # ---- Azioni creazione ----
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
        self.btn_create_sw.setVisible(bool(prt_asm_no_file))
        self.btn_create_file.setVisible(bool(prt_asm_no_file))
        self.grp_actions.setVisible(bool(prt_asm_no_file))

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
                    # Apertura shell di fallback
                    subprocess.Popen(["cmd", "/c", "start", "", str(target)])
            except Exception:
                QMessageBox.critical(self, "Errore apertura SolidWorks", str(e))

    # ------------------------------------------------------------------
    #  Thumbnail
    # ------------------------------------------------------------------
    def _load_thumb_to(self, doc: dict, lbl: QLabel):
        """Carica thumbnail in un QLabel specifico (con fallback icona tipo)."""
        thumb_path = self._get_thumbnail_path(doc)
        _sz = lbl.width() or 110
        if thumb_path and thumb_path.exists():
            pixmap = QPixmap(str(thumb_path))
            if not pixmap.isNull():
                scaled = pixmap.scaled(
                    _sz, _sz,
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
        self.lbl_thumb.setPixmap(QPixmap())
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

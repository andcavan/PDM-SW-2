# =============================================================================
#  ui/sw_config_dialog.py  –  Configurazione SolidWorks (template e .reg)
# =============================================================================
from __future__ import annotations
import subprocess
from pathlib import Path

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QGroupBox, QFormLayout, QFileDialog,
    QMessageBox, QTabWidget, QWidget, QTextEdit
)
from PyQt6.QtCore import Qt

from config import load_local_config, save_local_config


# Chiavi usate in local_config.json
CFG_TPL_PRT    = "sw_template_prt"
CFG_TPL_ASM    = "sw_template_asm"
CFG_TPL_DRW    = "sw_template_drw"
CFG_REG_FILE   = "sw_reg_file"
CFG_WORKSPACE  = "sw_workspace"


class SWConfigDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Configurazione SolidWorks")
        self.setMinimumWidth(620)
        self._cfg = load_local_config()
        self._build_ui()
        self._load()

    # ==================================================================
    # UI
    # ==================================================================
    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        lbl = QLabel(
            "Configura i template SolidWorks e le impostazioni locali di questa workstation.\n"
            "Le impostazioni sono salvate su questo PC (non nella cartella condivisa)."
        )
        lbl.setObjectName("subtitle_label")
        lbl.setWordWrap(True)
        layout.addWidget(lbl)

        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_templates_tab(), "Template")
        self.tabs.addTab(self._build_sw_tab(),        "SolidWorks")
        self.tabs.addTab(self._build_workspace_tab(), "Workspace locale")
        layout.addWidget(self.tabs)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_save = QPushButton("Salva")
        btn_save.setObjectName("btn_primary")
        btn_save.clicked.connect(self._save)
        btn_close = QPushButton("Chiudi")
        btn_close.clicked.connect(self.accept)
        btn_row.addWidget(btn_close)
        btn_row.addWidget(btn_save)
        layout.addLayout(btn_row)

    # ------------------------------------------------------------------
    def _build_templates_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)

        grp = QGroupBox("File template SolidWorks")
        form = QFormLayout(grp)

        self.txt_tpl_prt = self._path_row(form, "Template Parte (.prtdot):",
                                          "Template SolidWorks Parte (*.prtdot *.PRTDOT)")
        self.txt_tpl_asm = self._path_row(form, "Template Assieme (.asmdot):",
                                          "Template SolidWorks Assieme (*.asmdot *.ASMDOT)")
        self.txt_tpl_drw = self._path_row(form, "Template Disegno (.drwdot):",
                                          "Template SolidWorks Disegno (*.drwdot *.DRWDOT)")

        info = QLabel(
            "I template vengono copiati e rinominati con il codice PDM "
            "quando si usa il pulsante 'Crea in SW' nel dettaglio documento."
        )
        info.setWordWrap(True)
        info.setObjectName("subtitle_label")

        layout.addWidget(grp)
        layout.addWidget(info)
        layout.addStretch()
        return w

    # ------------------------------------------------------------------
    def _build_sw_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)

        grp_reg = QGroupBox("File di registro SolidWorks (.reg)")
        form_reg = QFormLayout(grp_reg)

        reg_row = QHBoxLayout()
        self.txt_reg = QLineEdit()
        self.txt_reg.setPlaceholderText("C:\\...\\solidworks_settings.reg")
        self.txt_reg.setReadOnly(True)
        btn_browse = QPushButton("Sfoglia…")
        btn_browse.clicked.connect(lambda: self._browse_path(
            self.txt_reg, "File registro (*.reg)"
        ))
        btn_apply_reg = QPushButton("Applica .reg")
        btn_apply_reg.setObjectName("btn_warning")
        btn_apply_reg.setToolTip(
            "Importa il file .reg nel registro di Windows (richiede privilegi amministrativi)"
        )
        btn_apply_reg.clicked.connect(self._apply_reg)
        reg_row.addWidget(self.txt_reg, stretch=1)
        reg_row.addWidget(btn_browse)
        reg_row.addWidget(btn_apply_reg)
        form_reg.addRow(reg_row)

        note_reg = QLabel(
            "⚠  L'applicazione del file .reg modifica il registro di sistema e "
            "potrebbe richiedere il riavvio di SolidWorks."
        )
        note_reg.setWordWrap(True)
        note_reg.setObjectName("subtitle_label")

        layout.addWidget(grp_reg)
        layout.addWidget(note_reg)
        layout.addStretch()
        return w

    # ------------------------------------------------------------------
    def _build_workspace_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)

        grp = QGroupBox("Workspace locale")
        form = QFormLayout(grp)

        ws_row = QHBoxLayout()
        self.txt_workspace = QLineEdit()
        self.txt_workspace.setPlaceholderText("C:\\Lavoro\\PDM_Workspace")
        btn_browse_ws = QPushButton("Sfoglia…")
        btn_browse_ws.clicked.connect(self._browse_workspace)
        ws_row.addWidget(self.txt_workspace, stretch=1)
        ws_row.addWidget(btn_browse_ws)
        form.addRow("Cartella workspace:", ws_row)

        info = QLabel(
            "Cartella locale dove vengono copiati i file in checkout.\n"
            "Ogni utente può avere una cartella diversa su questa workstation."
        )
        info.setWordWrap(True)
        info.setObjectName("subtitle_label")

        layout.addWidget(grp)
        layout.addWidget(info)
        layout.addStretch()
        return w

    # ==================================================================
    # Helpers UI
    # ==================================================================
    def _path_row(self, form: QFormLayout, label: str, file_filter: str) -> QLineEdit:
        row = QHBoxLayout()
        txt = QLineEdit()
        txt.setReadOnly(True)
        btn = QPushButton("Sfoglia…")
        btn.clicked.connect(lambda checked=False, t=txt, f=file_filter:
                            self._browse_path(t, f))
        clear_btn = QPushButton("✕")
        clear_btn.setMaximumWidth(28)
        clear_btn.setToolTip("Rimuovi")
        clear_btn.clicked.connect(lambda checked=False, t=txt: t.clear())
        row.addWidget(txt, stretch=1)
        row.addWidget(btn)
        row.addWidget(clear_btn)
        form.addRow(label, row)
        return txt

    def _browse_path(self, target: QLineEdit, file_filter: str):
        path, _ = QFileDialog.getOpenFileName(self, "Seleziona file", "", file_filter)
        if path:
            target.setText(path)

    def _browse_workspace(self):
        folder = QFileDialog.getExistingDirectory(self, "Seleziona cartella workspace")
        if folder:
            self.txt_workspace.setText(folder)

    # ==================================================================
    # Load / Save
    # ==================================================================
    def _load(self):
        self.txt_tpl_prt.setText(self._cfg.get(CFG_TPL_PRT, ""))
        self.txt_tpl_asm.setText(self._cfg.get(CFG_TPL_ASM, ""))
        self.txt_tpl_drw.setText(self._cfg.get(CFG_TPL_DRW, ""))
        self.txt_reg.setText(self._cfg.get(CFG_REG_FILE, ""))
        self.txt_workspace.setText(self._cfg.get(CFG_WORKSPACE, ""))

    def _save(self):
        self._cfg[CFG_TPL_PRT]   = self.txt_tpl_prt.text().strip()
        self._cfg[CFG_TPL_ASM]   = self.txt_tpl_asm.text().strip()
        self._cfg[CFG_TPL_DRW]   = self.txt_tpl_drw.text().strip()
        self._cfg[CFG_REG_FILE]   = self.txt_reg.text().strip()
        self._cfg[CFG_WORKSPACE]  = self.txt_workspace.text().strip()
        save_local_config(self._cfg)
        QMessageBox.information(self, "OK", "Configurazione SolidWorks salvata")

    def _apply_reg(self):
        reg_path = self.txt_reg.text().strip()
        if not reg_path or not Path(reg_path).exists():
            QMessageBox.warning(self, "Errore", "File .reg non trovato")
            return
        r = QMessageBox.question(
            self, "Applica file .reg",
            f"Importare nel registro di sistema:\n{reg_path}\n\n"
            "Continuare?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if r != QMessageBox.StandardButton.Yes:
            return
        try:
            result = subprocess.run(
                ["reg", "import", reg_path],
                capture_output=True, text=True, timeout=30
            )
            if result.returncode == 0:
                QMessageBox.information(
                    self, "OK",
                    "File .reg importato con successo.\n"
                    "Riavviare SolidWorks per applicare le impostazioni."
                )
            else:
                QMessageBox.critical(
                    self, "Errore",
                    f"Importazione fallita:\n{result.stderr or result.stdout}"
                )
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

    # ==================================================================
    # Metodi statici di utilità (usati da document_dialog)
    # ==================================================================
    @staticmethod
    def get_template(doc_type: str) -> Path | None:
        """Ritorna il path del template per il tipo documento, o None."""
        cfg = load_local_config()
        map_ = {
            "Parte":    cfg.get(CFG_TPL_PRT, ""),
            "Assieme":  cfg.get(CFG_TPL_ASM, ""),
            "Disegno":  cfg.get(CFG_TPL_DRW, ""),
        }
        p = map_.get(doc_type, "")
        return Path(p) if p else None

    @staticmethod
    def get_workspace() -> Path | None:
        """Ritorna la workspace locale configurata, o None."""
        cfg = load_local_config()
        p = cfg.get(CFG_WORKSPACE, "")
        return Path(p) if p else None

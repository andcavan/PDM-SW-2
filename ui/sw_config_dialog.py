# =============================================================================
#  ui/sw_config_dialog.py  –  Configurazione SolidWorks (template, exe, .reg)
# =============================================================================
from __future__ import annotations
import subprocess
from pathlib import Path

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QGroupBox, QFormLayout, QFileDialog,
    QMessageBox, QTabWidget, QWidget, QTextEdit, QCheckBox,
    QRadioButton, QButtonGroup
)
from PyQt6.QtCore import Qt

from config import load_local_config, save_local_config
from core.reg_manager import (
    RestoreOptions, import_reg, detect_solidworks_exe, detect_edrawings_exe,
)


# Chiavi usate in local_config.json
CFG_TPL_PRT    = "sw_template_prt"
CFG_TPL_ASM    = "sw_template_asm"
CFG_TPL_DRW    = "sw_template_drw"
CFG_REG_FILE   = "sw_reg_file"
CFG_WORKSPACE  = "sw_workspace"
CFG_SW_EXE     = "sw_exe_path"
CFG_EDRAW_EXE  = "edrawings_exe_path"


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

        # ---- Eseguibili ----
        grp_exe = QGroupBox("Eseguibili")
        form_exe = QFormLayout(grp_exe)

        # SolidWorks exe
        sw_row = QHBoxLayout()
        self.txt_sw_exe = QLineEdit()
        self.txt_sw_exe.setPlaceholderText("C:\\...\\SLDWORKS.exe")
        self.txt_sw_exe.setReadOnly(True)
        btn_sw_browse = QPushButton("Sfoglia…")
        btn_sw_browse.clicked.connect(lambda: self._browse_path(
            self.txt_sw_exe, "SolidWorks (SLDWORKS.exe)"
        ))
        btn_sw_detect = QPushButton("Rileva")
        btn_sw_detect.clicked.connect(self._detect_sw)
        sw_row.addWidget(self.txt_sw_exe, stretch=1)
        sw_row.addWidget(btn_sw_browse)
        sw_row.addWidget(btn_sw_detect)
        form_exe.addRow("SolidWorks:", sw_row)

        # eDrawings exe
        ed_row = QHBoxLayout()
        self.txt_edraw_exe = QLineEdit()
        self.txt_edraw_exe.setPlaceholderText("C:\\...\\EModelView.exe")
        self.txt_edraw_exe.setReadOnly(True)
        btn_ed_browse = QPushButton("Sfoglia…")
        btn_ed_browse.clicked.connect(lambda: self._browse_path(
            self.txt_edraw_exe, "eDrawings (EModelView.exe *.exe)"
        ))
        btn_ed_detect = QPushButton("Rileva")
        btn_ed_detect.clicked.connect(self._detect_edraw)
        ed_row.addWidget(self.txt_edraw_exe, stretch=1)
        ed_row.addWidget(btn_ed_browse)
        ed_row.addWidget(btn_ed_detect)
        form_exe.addRow("eDrawings:", ed_row)

        layout.addWidget(grp_exe)

        # ---- File di registro ----
        grp_reg = QGroupBox("File di registro SolidWorks (.reg / .sldreg)")
        form_reg = QFormLayout(grp_reg)

        reg_row = QHBoxLayout()
        self.txt_reg = QLineEdit()
        self.txt_reg.setPlaceholderText("C:\\...\\solidworks_settings.reg")
        self.txt_reg.setReadOnly(True)
        btn_browse = QPushButton("Sfoglia…")
        btn_browse.clicked.connect(lambda: self._browse_path(
            self.txt_reg, "File registro (*.reg *.sldreg)"
        ))
        reg_row.addWidget(self.txt_reg, stretch=1)
        reg_row.addWidget(btn_browse)
        form_reg.addRow("File:", reg_row)

        layout.addWidget(grp_reg)

        # ---- Cosa ripristinare ----
        grp_opts = QGroupBox("Cosa ripristinare")
        opts_layout = QVBoxLayout(grp_opts)

        self.chk_system = QCheckBox("Opzioni del sistema")
        self.chk_system.setChecked(True)
        opts_layout.addWidget(self.chk_system)

        self.chk_toolbar = QCheckBox("Layout barre strumenti")
        self.chk_toolbar.setChecked(True)
        self.chk_toolbar.toggled.connect(self._sync_toolbar_radios)
        opts_layout.addWidget(self.chk_toolbar)

        # Radio toolbar mode (indentate)
        tb_group = QHBoxLayout()
        tb_group.setContentsMargins(24, 0, 0, 0)
        self.rad_tb_all = QRadioButton("Tutte le barre + CommandManager")
        self.rad_tb_all.setChecked(True)
        self.rad_tb_macro = QRadioButton("Solo barra macro")
        self._tb_group = QButtonGroup(self)
        self._tb_group.addButton(self.rad_tb_all)
        self._tb_group.addButton(self.rad_tb_macro)
        tb_group.addWidget(self.rad_tb_all)
        tb_group.addWidget(self.rad_tb_macro)
        tb_group.addStretch()
        opts_layout.addLayout(tb_group)

        self.chk_keys = QCheckBox("Tasti rapidi da tastiera")
        self.chk_keys.setChecked(True)
        opts_layout.addWidget(self.chk_keys)

        self.chk_mouse = QCheckBox("Gesti del mouse")
        self.chk_mouse.setChecked(True)
        opts_layout.addWidget(self.chk_mouse)

        self.chk_menu = QCheckBox("Personalizzazioni menu")
        self.chk_menu.setChecked(True)
        opts_layout.addWidget(self.chk_menu)

        self.chk_views = QCheckBox("Viste salvate")
        self.chk_views.setChecked(True)
        opts_layout.addWidget(self.chk_views)

        self.chk_cleanup = QCheckBox("Pulisci chiavi selezionate prima dell'import")
        self.chk_cleanup.setChecked(True)
        self.chk_cleanup.setToolTip(
            "Cancella le chiavi di registro corrispondenti prima di importare "
            "(evita residui da configurazioni precedenti)"
        )
        opts_layout.addWidget(self.chk_cleanup)

        layout.addWidget(grp_opts)

        # ---- Bottone applica ----
        btn_apply = QPushButton("Applica configurazione registro")
        btn_apply.setObjectName("btn_warning")
        btn_apply.setToolTip(
            "Importa le sezioni selezionate nel registro di Windows"
        )
        btn_apply.clicked.connect(self._apply_reg)
        layout.addWidget(btn_apply)

        note = QLabel(
            "⚠  L'applicazione modifica il registro di sistema.\n"
            "Chiudere SolidWorks prima di applicare e riavviarlo dopo."
        )
        note.setWordWrap(True)
        note.setObjectName("subtitle_label")
        layout.addWidget(note)

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
        self.txt_sw_exe.setText(self._cfg.get(CFG_SW_EXE, ""))
        self.txt_edraw_exe.setText(self._cfg.get(CFG_EDRAW_EXE, ""))

    def _save(self):
        self._cfg[CFG_TPL_PRT]    = self.txt_tpl_prt.text().strip()
        self._cfg[CFG_TPL_ASM]    = self.txt_tpl_asm.text().strip()
        self._cfg[CFG_TPL_DRW]    = self.txt_tpl_drw.text().strip()
        self._cfg[CFG_REG_FILE]   = self.txt_reg.text().strip()
        self._cfg[CFG_WORKSPACE]  = self.txt_workspace.text().strip()
        self._cfg[CFG_SW_EXE]     = self.txt_sw_exe.text().strip()
        self._cfg[CFG_EDRAW_EXE]  = self.txt_edraw_exe.text().strip()
        save_local_config(self._cfg)
        QMessageBox.information(self, "OK", "Configurazione SolidWorks salvata")

    def _apply_reg(self):
        reg_path = self.txt_reg.text().strip()
        if not reg_path or not Path(reg_path).exists():
            QMessageBox.warning(self, "Errore", "File .reg non trovato")
            return

        options = self._collect_restore_options()
        if not options.has_any_selection():
            QMessageBox.warning(
                self, "Nessuna selezione",
                "Seleziona almeno una voce da ripristinare."
            )
            return

        r = QMessageBox.question(
            self, "Applica configurazione registro",
            f"Importare nel registro di sistema:\n{reg_path}\n\n"
            f"Voci selezionate: {options.describe()}\n\n"
            "Continuare?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if r != QMessageBox.StandardButton.Yes:
            return

        try:
            ok, msg = import_reg(Path(reg_path), options)
            if ok:
                QMessageBox.information(
                    self, "OK",
                    f"{msg}\n\nRiavviare SolidWorks per applicare."
                )
            else:
                QMessageBox.critical(self, "Errore", f"Importazione fallita:\n{msg}")
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

    def _collect_restore_options(self) -> RestoreOptions:
        return RestoreOptions(
            system_options=self.chk_system.isChecked(),
            toolbar_layout=self.chk_toolbar.isChecked(),
            toolbar_mode="macro_only" if self.rad_tb_macro.isChecked() else "all",
            keyboard_shortcuts=self.chk_keys.isChecked(),
            mouse_gestures=self.chk_mouse.isChecked(),
            menu_customizations=self.chk_menu.isChecked(),
            saved_views=self.chk_views.isChecked(),
            cleanup_before_import=self.chk_cleanup.isChecked(),
        )

    def _sync_toolbar_radios(self, checked: bool):
        self.rad_tb_all.setEnabled(checked)
        self.rad_tb_macro.setEnabled(checked)

    def _detect_sw(self):
        found = detect_solidworks_exe()
        if found:
            self.txt_sw_exe.setText(str(found))
            QMessageBox.information(self, "Rilevato", f"SolidWorks trovato:\n{found}")
        else:
            QMessageBox.warning(
                self, "Non trovato",
                "SLDWORKS.exe non trovato automaticamente.\n"
                "Usa 'Sfoglia' per selezionarlo manualmente."
            )

    def _detect_edraw(self):
        found = detect_edrawings_exe()
        if found:
            self.txt_edraw_exe.setText(str(found))
            QMessageBox.information(self, "Rilevato", f"eDrawings trovato:\n{found}")
        else:
            QMessageBox.warning(
                self, "Non trovato",
                "EModelView.exe non trovato automaticamente.\n"
                "Usa 'Sfoglia' per selezionarlo manualmente."
            )

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

    @staticmethod
    def get_solidworks_exe() -> Path | None:
        """Ritorna il path dell'eseguibile SolidWorks configurato, o None."""
        cfg = load_local_config()
        p = cfg.get(CFG_SW_EXE, "")
        return Path(p) if p and Path(p).is_file() else None

    @staticmethod
    def get_edrawings_exe() -> Path | None:
        """Ritorna il path dell'eseguibile eDrawings configurato, o None."""
        cfg = load_local_config()
        p = cfg.get(CFG_EDRAW_EXE, "")
        return Path(p) if p and Path(p).is_file() else None

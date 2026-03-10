# =============================================================================
#  ui/coding_schema_dialog.py  –  Configurazione schema di codifica
# =============================================================================
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QGroupBox, QFormLayout, QSpinBox,
    QMessageBox, QScrollArea, QWidget, QFrame
)
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QFont

from core.coding_config import CodingSchemeConfig, LevelConfig
from ui.session import session


class CodingSchemaDialog(QDialog):
    """Dialog per configurare lo schema di codifica. Solo Amministratori."""

    _PREVIEW_MACH = "ABC"
    _PREVIEW_GRP  = "COMP"
    _PREVIEW_VER  = 1

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Schema di Codifica")
        self.setMinimumSize(640, 780)
        self._cfg = session.coding.get_scheme_config()
        self._preview_timer = QTimer(self)
        self._preview_timer.setSingleShot(True)
        self._preview_timer.setInterval(200)
        self._preview_timer.timeout.connect(self._refresh_preview)
        self._build_ui()
        self._load_cfg_to_ui()
        self._refresh_preview()

    # ==================================================================
    # UI
    # ==================================================================
    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setSpacing(10)
        root.setContentsMargins(16, 16, 16, 16)

        info = QLabel(
            "Configura il formato dei codici per ogni livello gerarchico.\n"
            "Le modifiche si applicano solo ai nuovi codici."
        )
        info.setObjectName("subtitle_label")
        info.setWordWrap(True)
        root.addWidget(info)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        inner = QWidget()
        inner_layout = QVBoxLayout(inner)
        inner_layout.setSpacing(10)

        # ── Nome schema ───────────────────────────────────────────────
        name_grp = QGroupBox("Nome schema")
        name_form = QFormLayout(name_grp)
        self.txt_name = QLineEdit()
        self.txt_name.setPlaceholderText("Es. Standard Gerarchico")
        self.txt_name.textChanged.connect(self._schedule_preview)
        name_form.addRow("Nome:", self.txt_name)
        inner_layout.addWidget(name_grp)

        # ── Formato codici Macchina e Gruppo ──────────────────────────
        from PyQt6.QtWidgets import QComboBox
        mach_grp = QGroupBox("Formato codici Macchina e Gruppo")
        mach_form = QFormLayout(mach_grp)

        mach_row = QHBoxLayout()
        self.cmb_mach_type = QComboBox()
        self.cmb_mach_type.addItems(["ALPHA (lettere A-Z)", "NUM (cifre 0-9)"])
        self.cmb_mach_type.currentIndexChanged.connect(self._schedule_preview)
        self.spn_mach_len = QSpinBox()
        self.spn_mach_len.setRange(2, 8)
        self.spn_mach_len.setSuffix(" caratteri")
        self.spn_mach_len.valueChanged.connect(self._schedule_preview)
        mach_row.addWidget(self.cmb_mach_type)
        mach_row.addWidget(self.spn_mach_len)
        mach_row.addStretch()
        mach_form.addRow("Codice macchina:", mach_row)

        grp_row = QHBoxLayout()
        self.cmb_grp_type = QComboBox()
        self.cmb_grp_type.addItems(["ALPHA (lettere A-Z)", "NUM (cifre 0-9)"])
        self.cmb_grp_type.currentIndexChanged.connect(self._schedule_preview)
        self.spn_grp_len = QSpinBox()
        self.spn_grp_len.setRange(2, 8)
        self.spn_grp_len.setSuffix(" caratteri")
        self.spn_grp_len.valueChanged.connect(self._schedule_preview)
        grp_row.addWidget(self.cmb_grp_type)
        grp_row.addWidget(self.spn_grp_len)
        grp_row.addStretch()
        mach_form.addRow("Codice gruppo:", grp_row)
        inner_layout.addWidget(mach_grp)

        # ── Variabili template ────────────────────────────────────────
        vars_grp = QGroupBox("Variabili disponibili nel template")
        vars_lyt = QVBoxLayout(vars_grp)
        vars_lbl = QLabel(
            "<b>{MACH}</b> → codice macchina &nbsp;&nbsp; "
            "<b>{GRP}</b> → codice gruppo<br>"
            "<b>{VER:N}</b> → versione N cifre &nbsp;&nbsp; "
            "<b>{NUM:N}</b> → numero sequenziale N cifre"
        )
        vars_lbl.setTextFormat(Qt.TextFormat.RichText)
        vars_lbl.setWordWrap(True)
        vars_lyt.addWidget(vars_lbl)
        inner_layout.addWidget(vars_grp)

        # ── LIV0 e LIV1 ──────────────────────────────────────────────
        self._liv0_w = self._build_tmpl_group("LIV0 — Macchina (ASM)")
        self._liv1_w = self._build_tmpl_group("LIV1 — Gruppo (ASM)")
        inner_layout.addWidget(self._liv0_w["grp"])
        inner_layout.addWidget(self._liv1_w["grp"])

        # ── LIV2/1 e LIV2/2 ──────────────────────────────────────────
        self._liv2_1_w = self._build_num_group("LIV2/1 — Sottogruppo (ASM)")
        self._liv2_2_w = self._build_num_group("LIV2/2 — Parte (PRT)")
        inner_layout.addWidget(self._liv2_1_w["grp"])
        inner_layout.addWidget(self._liv2_2_w["grp"])

        # ── Warning collisione ────────────────────────────────────────
        coll_grp = QGroupBox("Warning collisione LIV2")
        coll_form = QFormLayout(coll_grp)
        self.spn_collision = QSpinBox()
        self.spn_collision.setRange(1, 100000)
        self.spn_collision.setValue(500)
        self.spn_collision.setSuffix(" codici rimasti")
        self.spn_collision.valueChanged.connect(self._schedule_preview)
        coll_form.addRow("Avvisa sotto:", self.spn_collision)
        inner_layout.addWidget(coll_grp)

        inner_layout.addStretch()
        scroll.setWidget(inner)
        root.addWidget(scroll, stretch=1)

        # ── Anteprima ─────────────────────────────────────────────────
        prev_grp = QGroupBox("Anteprima (MACH=ABC, GRP=COMP, primo codice)")
        prev_lyt = QVBoxLayout(prev_grp)
        self.lbl_preview = QLabel()
        self.lbl_preview.setTextFormat(Qt.TextFormat.RichText)
        self.lbl_preview.setFont(QFont("Consolas", 10))
        self.lbl_preview.setWordWrap(True)
        prev_lyt.addWidget(self.lbl_preview)
        root.addWidget(prev_grp)

        # ── Pulsanti ──────────────────────────────────────────────────
        btn_row = QHBoxLayout()
        btn_default = QPushButton("Ripristina Default")
        btn_default.clicked.connect(self._restore_default)
        btn_save = QPushButton("Salva Schema")
        btn_save.setObjectName("btn_primary")
        btn_save.clicked.connect(self._save)
        btn_close = QPushButton("Chiudi")
        btn_close.clicked.connect(self.reject)
        btn_row.addWidget(btn_default)
        btn_row.addStretch()
        btn_row.addWidget(btn_close)
        btn_row.addWidget(btn_save)
        root.addLayout(btn_row)

    def _build_tmpl_group(self, title: str) -> dict:
        """Gruppo con template, prefisso, suffisso (LIV0, LIV1)."""
        grp = QGroupBox(title)
        form = QFormLayout(grp)

        txt_tmpl = QLineEdit()
        txt_tmpl.textChanged.connect(self._schedule_preview)
        form.addRow("Template:", txt_tmpl)

        txt_warn = QLabel()
        txt_warn.setStyleSheet("color: orange; font-size: 11px;")
        txt_warn.setVisible(False)
        form.addRow("", txt_warn)

        txt_prefix = QLineEdit()
        txt_prefix.setMaximumWidth(120)
        txt_prefix.setPlaceholderText("(vuoto)")
        txt_prefix.textChanged.connect(self._schedule_preview)
        form.addRow("Prefisso:", txt_prefix)

        txt_suffix = QLineEdit()
        txt_suffix.setMaximumWidth(120)
        txt_suffix.setPlaceholderText("(vuoto)")
        txt_suffix.textChanged.connect(self._schedule_preview)
        form.addRow("Suffisso:", txt_suffix)

        return {"grp": grp, "template": txt_tmpl, "warn": txt_warn,
                "prefix": txt_prefix, "suffix": txt_suffix}

    def _build_num_group(self, title: str) -> dict:
        """Gruppo con template, prefisso, suffisso e range inizio/fine (LIV2)."""
        grp = QGroupBox(title)
        form = QFormLayout(grp)

        txt_tmpl = QLineEdit()
        txt_tmpl.textChanged.connect(self._schedule_preview)
        form.addRow("Template:", txt_tmpl)

        txt_warn = QLabel()
        txt_warn.setStyleSheet("color: orange; font-size: 11px;")
        txt_warn.setVisible(False)
        form.addRow("", txt_warn)

        txt_prefix = QLineEdit()
        txt_prefix.setMaximumWidth(120)
        txt_prefix.setPlaceholderText("(vuoto)")
        txt_prefix.textChanged.connect(self._schedule_preview)
        form.addRow("Prefisso:", txt_prefix)

        txt_suffix = QLineEdit()
        txt_suffix.setMaximumWidth(120)
        txt_suffix.setPlaceholderText("(vuoto)")
        txt_suffix.textChanged.connect(self._schedule_preview)
        form.addRow("Suffisso:", txt_suffix)

        # Range: inizio e fine su una riga + indicatore direzione auto
        range_row = QHBoxLayout()
        spn_inizio = QSpinBox()
        spn_inizio.setRange(0, 999999)
        spn_inizio.valueChanged.connect(self._schedule_preview)
        spn_fine = QSpinBox()
        spn_fine.setRange(0, 999999)
        spn_fine.valueChanged.connect(self._schedule_preview)
        lbl_dir = QLabel("↑")   # aggiornata live in _refresh_preview
        lbl_dir.setStyleSheet("font-size: 14px; font-weight: bold; color: #0080ff;")
        range_row.addWidget(QLabel("Inizio:"))
        range_row.addWidget(spn_inizio)
        range_row.addWidget(QLabel("  Fine:"))
        range_row.addWidget(spn_fine)
        range_row.addWidget(lbl_dir)
        range_row.addStretch()
        form.addRow("Range:", range_row)

        return {"grp": grp, "template": txt_tmpl, "warn": txt_warn,
                "prefix": txt_prefix, "suffix": txt_suffix,
                "spn_inizio": spn_inizio, "spn_fine": spn_fine,
                "lbl_dir": lbl_dir}

    # ==================================================================
    # Caricamento / lettura da UI
    # ==================================================================
    def _load_cfg_to_ui(self):
        cfg = self._cfg
        self.txt_name.setText(cfg.name)

        self.cmb_mach_type.setCurrentIndex(0 if cfg.mach_code_type == "ALPHA" else 1)
        self.spn_mach_len.setValue(cfg.mach_code_length)
        self.cmb_grp_type.setCurrentIndex(0 if cfg.grp_code_type == "ALPHA" else 1)
        self.spn_grp_len.setValue(cfg.grp_code_length)

        self._load_tmpl(self._liv0_w, cfg.liv0)
        self._load_tmpl(self._liv1_w, cfg.liv1)
        self._load_num(self._liv2_1_w, cfg.liv2_1)
        self._load_num(self._liv2_2_w, cfg.liv2_2)

        self.spn_collision.setValue(cfg.liv2_1.collision_threshold)

    def _load_tmpl(self, w: dict, lc: LevelConfig):
        w["template"].setText(lc.template)
        w["prefix"].setText(lc.prefix)
        w["suffix"].setText(lc.suffix)

    def _load_num(self, w: dict, lc: LevelConfig):
        w["template"].setText(lc.template)
        w["prefix"].setText(lc.prefix)
        w["suffix"].setText(lc.suffix)
        # inizio = num_start, fine = num_max (asc) o num_min (desc)
        w["spn_inizio"].setValue(lc.num_start)
        w["spn_fine"].setValue(lc.num_max if lc.num_dir == "asc" else lc.num_min)
        self._update_dir_label(w)

    @staticmethod
    def _update_dir_label(w: dict):
        inizio = w["spn_inizio"].value()
        fine   = w["spn_fine"].value()
        if inizio < fine:
            w["lbl_dir"].setText("↑")
            w["lbl_dir"].setToolTip("Ascendente: il contatore sale da Inizio a Fine")
        elif inizio > fine:
            w["lbl_dir"].setText("↓")
            w["lbl_dir"].setToolTip("Discendente: il contatore scende da Inizio a Fine")
        else:
            w["lbl_dir"].setText("=")
            w["lbl_dir"].setToolTip("Inizio e Fine uguali — range non valido")

    @staticmethod
    def _lc_from_range(inizio: int, fine: int) -> tuple:
        """Ritorna (num_start, num_max, num_min, num_dir) dai valori inizio/fine."""
        if inizio <= fine:
            return inizio, fine, inizio, "asc"
        else:
            return inizio, fine, fine, "desc"

    def _read_tmpl(self, w: dict, default: LevelConfig) -> LevelConfig:
        return LevelConfig(
            template=w["template"].text().strip() or default.template,
            prefix=w["prefix"].text(),
            suffix=w["suffix"].text(),
            num_start=default.num_start, num_max=default.num_max,
            num_min=default.num_min, num_dir=default.num_dir,
            collision_threshold=self.spn_collision.value(),
        )

    def _read_num(self, w: dict, default: LevelConfig) -> LevelConfig:
        inizio = w["spn_inizio"].value()
        fine   = w["spn_fine"].value()
        num_start, num_max, num_min, num_dir = self._lc_from_range(inizio, fine)
        return LevelConfig(
            template=w["template"].text().strip() or default.template,
            prefix=w["prefix"].text(),
            suffix=w["suffix"].text(),
            num_start=num_start, num_max=num_max,
            num_min=num_min, num_dir=num_dir,
            collision_threshold=self.spn_collision.value(),
        )

    def _build_cfg_from_ui(self) -> CodingSchemeConfig:
        from PyQt6.QtWidgets import QComboBox
        default = CodingSchemeConfig.default()
        return CodingSchemeConfig(
            name=self.txt_name.text().strip() or "Schema Personalizzato",
            mach_code_type="ALPHA" if self.cmb_mach_type.currentIndex() == 0 else "NUM",
            mach_code_length=self.spn_mach_len.value(),
            grp_code_type="ALPHA" if self.cmb_grp_type.currentIndex() == 0 else "NUM",
            grp_code_length=self.spn_grp_len.value(),
            liv0=self._read_tmpl(self._liv0_w, default.liv0),
            liv1=self._read_tmpl(self._liv1_w, default.liv1),
            liv2_1=self._read_num(self._liv2_1_w, default.liv2_1),
            liv2_2=self._read_num(self._liv2_2_w, default.liv2_2),
        )

    # ==================================================================
    # Preview
    # ==================================================================
    def _schedule_preview(self):
        self._preview_timer.start()
        for w in (self._liv0_w, self._liv1_w, self._liv2_1_w, self._liv2_2_w):
            tmpl = w["template"].text().strip()
            unknown = CodingSchemeConfig.validate_template(tmpl)
            if unknown:
                w["warn"].setText(f"⚠ Variabili non riconosciute: {', '.join(unknown)}")
                w["warn"].setVisible(True)
            else:
                w["warn"].setVisible(False)
        # aggiorna frecce direzione
        self._update_dir_label(self._liv2_1_w)
        self._update_dir_label(self._liv2_2_w)

    def _refresh_preview(self):
        try:
            cfg = self._build_cfg_from_ui()
        except Exception:
            return

        m, g, v = self._PREVIEW_MACH, self._PREVIEW_GRP, self._PREVIEW_VER

        def render(lc, **kw) -> str:
            try:
                return cfg.render(lc, **kw)
            except Exception:
                return "<i style='color:red'>errore template</i>"

        lc1, lc2 = cfg.liv2_1, cfg.liv2_2
        fine1 = lc1.num_max if lc1.num_dir == "asc" else lc1.num_min
        fine2 = lc2.num_max if lc2.num_dir == "asc" else lc2.num_min
        d1 = "↑" if lc1.num_dir == "asc" else "↓"
        d2 = "↑" if lc2.num_dir == "asc" else "↓"

        lines = [
            f"<b>Macchina:</b> {cfg.mach_code_type}, {cfg.mach_code_length} car. "
            f"&nbsp;&nbsp; <b>Gruppo:</b> {cfg.grp_code_type}, {cfg.grp_code_length} car.",
            "",
            f"<b>LIV0:</b>   {render(cfg.liv0, mach=m, ver=v)}",
            f"<b>LIV1:</b>   {render(cfg.liv1, mach=m, grp=g, ver=v)}",
            f"<b>LIV2/1 (ASM):</b>   {render(lc1, mach=m, grp=g, num=lc1.num_start)}"
            f"  <span style='color:gray;font-size:10px'>{d1} {lc1.num_start}→{fine1}</span>",
            f"<b>LIV2/2 (PRT):</b>   {render(lc2, mach=m, grp=g, num=lc2.num_start)}"
            f"  <span style='color:gray;font-size:10px'>{d2} {lc2.num_start}→{fine2}</span>",
        ]
        self.lbl_preview.setText("<br>".join(lines))

    # ==================================================================
    # Azioni pulsanti
    # ==================================================================
    def _save(self):
        cfg = self._build_cfg_from_ui()

        for attr, label in [
            ("liv0", "LIV0"), ("liv1", "LIV1"),
            ("liv2_1", "LIV2/1"), ("liv2_2", "LIV2/2"),
        ]:
            lc: LevelConfig = getattr(cfg, attr)
            if not lc.template.strip():
                QMessageBox.warning(self, "Template mancante",
                                    f"Il template per {label} non può essere vuoto.")
                return
            unknown = CodingSchemeConfig.validate_template(lc.template)
            if unknown:
                r = QMessageBox.question(
                    self, "Variabili non riconosciute",
                    f"Template {label} contiene variabili non riconosciute: "
                    f"{', '.join(unknown)}.\n\nSalvare comunque?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                )
                if r != QMessageBox.StandardButton.Yes:
                    return

        # Valida range: inizio != fine
        for lc, label in [(cfg.liv2_1, "LIV2/1"), (cfg.liv2_2, "LIV2/2")]:
            fine = lc.num_max if lc.num_dir == "asc" else lc.num_min
            if lc.num_start == fine:
                QMessageBox.warning(self, f"Range {label} non valido",
                    f"{label}: Inizio e Fine non possono essere uguali.")
                return

        try:
            session.coding.save_scheme_config(cfg)
        except Exception as e:
            QMessageBox.critical(self, "Errore salvataggio", str(e))
            return

        QMessageBox.information(
            self, "Schema salvato",
            "Lo schema di codifica è stato salvato.\n\n"
            "I parametri si applicano solo ai nuovi codici.\n"
            "I codici già presenti nel database non vengono modificati."
        )
        self._cfg = cfg

    def _restore_default(self):
        r = QMessageBox.question(
            self, "Ripristina default",
            "Ripristinare tutti i valori allo schema standard?\n"
            "(Le modifiche non salvate andranno perse.)",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if r == QMessageBox.StandardButton.Yes:
            self._cfg = CodingSchemeConfig.default()
            self._load_cfg_to_ui()
            self._refresh_preview()

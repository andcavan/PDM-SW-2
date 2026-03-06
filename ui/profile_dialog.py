# =============================================================================
#  ui/profile_dialog.py  –  PDM Profile: gestione profili multi-ambiente
# =============================================================================
from __future__ import annotations

import shutil
from pathlib import Path

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QGroupBox, QFormLayout, QFileDialog,
    QMessageBox, QListWidget, QListWidgetItem, QSplitter,
    QComboBox, QRadioButton, QButtonGroup, QWidget,
    QInputDialog,
)
from PyQt6.QtCore import Qt, pyqtSignal

from config import (
    get_profile_names, get_active_profile_name, set_active_profile,
    load_profile, save_profile, delete_profile, rename_profile,
    copy_profile,
)


# ======================================================================
#  ProfileSelector — dialog compatto per scegliere profilo all'avvio
# ======================================================================

class ProfileSelector(QDialog):
    """Dialog semplice per selezionare un profilo all'avvio."""

    def __init__(self, profiles: list, active: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("PDM Profile")
        self.setMinimumWidth(380)
        self.setModal(True)
        self._selected = active
        self._build_ui(profiles, active)

    def _build_ui(self, profiles, active):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(24, 24, 24, 24)

        lbl = QLabel("📋  Seleziona profilo di lavoro:")
        lbl.setObjectName("subtitle_label")
        layout.addWidget(lbl)

        self.combo = QComboBox()
        self.combo.addItems(profiles)
        if active in profiles:
            self.combo.setCurrentText(active)
        self.combo.setMinimumHeight(30)
        layout.addWidget(self.combo)

        btn_row = QHBoxLayout()
        btn_manage = QPushButton("Gestisci profili…")
        btn_manage.clicked.connect(self._manage)
        btn_ok = QPushButton("Connetti")
        btn_ok.setObjectName("btn_primary")
        btn_ok.setDefault(True)
        btn_ok.clicked.connect(self._on_accept)
        btn_row.addWidget(btn_manage)
        btn_row.addStretch()
        btn_row.addWidget(btn_ok)
        layout.addLayout(btn_row)

    def _on_accept(self):
        self._selected = self.combo.currentText()
        self.accept()

    def _manage(self):
        dlg = ProfileDialog(self)
        dlg.exec()
        profiles = get_profile_names()
        self.combo.clear()
        self.combo.addItems(profiles)
        active = get_active_profile_name()
        if active in profiles:
            self.combo.setCurrentText(active)

    @property
    def selected_profile(self) -> str:
        return self._selected


# ======================================================================
#  CopyProfileDialog — scelta modalità copia
# ======================================================================

class CopyProfileDialog(QDialog):
    """Dialog per scegliere come copiare un profilo."""

    def __init__(self, src_profile: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Copia profilo")
        self.setMinimumWidth(480)
        self.setModal(True)
        self.dst_name = ""
        self.copy_data = False
        self.new_shared_root = ""
        self._build_ui(src_profile)

    def _build_ui(self, src_profile):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        layout.addWidget(QLabel(f"Copia del profilo: <b>{src_profile}</b>"))

        form = QFormLayout()
        self.txt_name = QLineEdit(f"{src_profile} (copia)")
        form.addRow("Nome nuovo profilo:", self.txt_name)
        layout.addLayout(form)

        # Modalità
        grp = QGroupBox("Cosa copiare")
        grp_layout = QVBoxLayout(grp)

        self.rad_config = QRadioButton("Solo configurazione")
        self.rad_config.setChecked(True)
        self.rad_config.setToolTip(
            "Copia le impostazioni (exe, template, workspace, .reg).\n"
            "Devi specificare un nuovo percorso condiviso."
        )
        grp_layout.addWidget(self.rad_config)

        self.rad_data = QRadioButton(
            "Configurazione + dati (archivio, thumbnail — escluso database)"
        )
        self.rad_data.setToolTip(
            "Copia le impostazioni e il contenuto della cartella condivisa\n"
            "(archivio, thumbnail, ecc.) escluso il database.\n"
            "Il nuovo ambiente parte con un database vuoto."
        )
        grp_layout.addWidget(self.rad_data)

        self._mode_group = QButtonGroup(self)
        self._mode_group.addButton(self.rad_config)
        self._mode_group.addButton(self.rad_data)

        layout.addWidget(grp)

        # Nuovo percorso condiviso
        form2 = QFormLayout()
        sr_row = QHBoxLayout()
        self.txt_shared = QLineEdit()
        self.txt_shared.setPlaceholderText("Percorso nuova cartella condivisa")
        btn_browse = QPushButton("Sfoglia…")
        btn_browse.clicked.connect(self._browse)
        sr_row.addWidget(self.txt_shared, stretch=1)
        sr_row.addWidget(btn_browse)
        form2.addRow("Percorso condiviso:", sr_row)
        layout.addLayout(form2)

        # Bottoni
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_cancel = QPushButton("Annulla")
        btn_cancel.clicked.connect(self.reject)
        btn_ok = QPushButton("Copia")
        btn_ok.setObjectName("btn_primary")
        btn_ok.clicked.connect(self._on_accept)
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_ok)
        layout.addLayout(btn_row)

    def _browse(self):
        path = QFileDialog.getExistingDirectory(
            self, "Seleziona cartella condivisa per il nuovo profilo"
        )
        if path:
            self.txt_shared.setText(path)

    def _on_accept(self):
        name = self.txt_name.text().strip()
        if not name:
            QMessageBox.warning(self, "Errore", "Inserire un nome per il profilo")
            return
        shared = self.txt_shared.text().strip()
        if not shared:
            QMessageBox.warning(
                self, "Errore",
                "Inserire il percorso condiviso per il nuovo profilo"
            )
            return
        self.dst_name = name
        self.copy_data = self.rad_data.isChecked()
        self.new_shared_root = shared
        self.accept()


# ======================================================================
#  ProfileDialog — CRUD completo profili
# ======================================================================

class ProfileDialog(QDialog):
    """Dialog per creare, modificare, copiare ed eliminare profili PDM."""

    profile_switched = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("PDM Profile – Gestione profili")
        self.setMinimumSize(720, 460)
        self.setModal(True)
        self._build_ui()
        self._refresh_list()

    # ------------------------------------------------------------------
    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(8)
        layout.setContentsMargins(12, 12, 12, 12)

        title = QLabel("📋  PDM Profile")
        title.setObjectName("title_label")
        layout.addWidget(title)

        splitter = QSplitter(Qt.Orientation.Horizontal)

        # ---- Sinistra: lista profili ----
        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 4, 0)

        left_layout.addWidget(QLabel("Profili disponibili:"))
        self.lst_profiles = QListWidget()
        self.lst_profiles.currentItemChanged.connect(self._on_selection_changed)
        left_layout.addWidget(self.lst_profiles)

        btn_grid = QHBoxLayout()
        btn_new = QPushButton("+ Nuovo")
        btn_new.clicked.connect(self._new_profile)
        btn_copy = QPushButton("⧉ Copia…")
        btn_copy.clicked.connect(self._copy_profile)
        btn_del = QPushButton("✕ Elimina")
        btn_del.clicked.connect(self._delete_profile)
        btn_grid.addWidget(btn_new)
        btn_grid.addWidget(btn_copy)
        btn_grid.addWidget(btn_del)
        left_layout.addLayout(btn_grid)

        # ---- Destra: dettagli profilo ----
        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(4, 0, 0, 0)

        grp = QGroupBox("Dettagli profilo")
        form = QFormLayout(grp)

        self.txt_name = QLineEdit()
        form.addRow("Nome:", self.txt_name)

        sr_row = QHBoxLayout()
        self.txt_shared = QLineEdit()
        self.txt_shared.setPlaceholderText(r"es. \\SERVER\PDM  oppure  Z:\PDM")
        btn_browse = QPushButton("Sfoglia…")
        btn_browse.clicked.connect(self._browse_shared)
        sr_row.addWidget(self.txt_shared, stretch=1)
        sr_row.addWidget(btn_browse)
        form.addRow("Percorso condiviso:", sr_row)

        # Summary impostazioni (sola lettura)
        self.lbl_sw_exe = QLabel("—")
        form.addRow("SolidWorks exe:", self.lbl_sw_exe)
        self.lbl_edraw = QLabel("—")
        form.addRow("eDrawings exe:", self.lbl_edraw)
        self.lbl_workspace = QLabel("—")
        form.addRow("Workspace:", self.lbl_workspace)
        self.lbl_templates = QLabel("—")
        form.addRow("Template:", self.lbl_templates)

        note = QLabel(
            "Per modificare exe, template e .reg usa:\n"
            "Strumenti → Configurazione SolidWorks (agisce sul profilo attivo)"
        )
        note.setWordWrap(True)
        note.setObjectName("subtitle_label")

        right_layout.addWidget(grp)
        right_layout.addWidget(note)

        # Bottoni azione profilo
        act_row = QHBoxLayout()
        btn_save = QPushButton("💾  Salva modifiche")
        btn_save.clicked.connect(self._save_profile)
        btn_activate = QPushButton("▶  Attiva profilo")
        btn_activate.setObjectName("btn_primary")
        btn_activate.clicked.connect(self._activate_profile)
        act_row.addWidget(btn_save)
        act_row.addWidget(btn_activate)
        right_layout.addLayout(act_row)

        right_layout.addStretch()

        splitter.addWidget(left)
        splitter.addWidget(right)
        splitter.setSizes([220, 500])
        layout.addWidget(splitter)

        # Chiudi
        close_row = QHBoxLayout()
        close_row.addStretch()
        btn_close = QPushButton("Chiudi")
        btn_close.clicked.connect(self.accept)
        close_row.addWidget(btn_close)
        layout.addLayout(close_row)

    # ------------------------------------------------------------------
    #  Refresh / selezione
    # ------------------------------------------------------------------
    def _refresh_list(self):
        self.lst_profiles.blockSignals(True)
        self.lst_profiles.clear()
        active = get_active_profile_name()
        for name in get_profile_names():
            display = f"● {name}" if name == active else name
            item = QListWidgetItem(display)
            item.setData(Qt.ItemDataRole.UserRole, name)
            if name == active:
                font = item.font()
                font.setBold(True)
                item.setFont(font)
            self.lst_profiles.addItem(item)
        self.lst_profiles.blockSignals(False)

    def _selected_name(self) -> str | None:
        item = self.lst_profiles.currentItem()
        return item.data(Qt.ItemDataRole.UserRole) if item else None

    def _on_selection_changed(self):
        name = self._selected_name()
        if not name:
            return
        data = load_profile(name)
        self.txt_name.setText(name)
        self.txt_shared.setText(data.get("shared_root", ""))
        self.lbl_sw_exe.setText(
            Path(data["sw_exe_path"]).name if data.get("sw_exe_path") else "—"
        )
        self.lbl_edraw.setText(
            Path(data["edrawings_exe_path"]).name
            if data.get("edrawings_exe_path") else "—"
        )
        self.lbl_workspace.setText(data.get("sw_workspace") or "—")
        tpls = []
        if data.get("sw_template_prt"):
            tpls.append("PRT")
        if data.get("sw_template_asm"):
            tpls.append("ASM")
        if data.get("sw_template_drw"):
            tpls.append("DRW")
        self.lbl_templates.setText(", ".join(tpls) if tpls else "—")

    # ------------------------------------------------------------------
    #  Azioni CRUD
    # ------------------------------------------------------------------
    def _browse_shared(self):
        path = QFileDialog.getExistingDirectory(
            self, "Seleziona cartella condivisa PDM"
        )
        if path:
            self.txt_shared.setText(path)

    def _new_profile(self):
        name, ok = QInputDialog.getText(
            self, "Nuovo profilo", "Nome del profilo:"
        )
        if not ok or not name.strip():
            return
        name = name.strip()
        if name in get_profile_names():
            QMessageBox.warning(self, "Errore", f"Il profilo '{name}' esiste già")
            return
        save_profile(name, {"shared_root": ""})
        self._refresh_list()
        # Seleziona il nuovo profilo
        for i in range(self.lst_profiles.count()):
            item = self.lst_profiles.item(i)
            if item and item.data(Qt.ItemDataRole.UserRole) == name:
                self.lst_profiles.setCurrentRow(i)
                break

    def _delete_profile(self):
        name = self._selected_name()
        if not name:
            return
        if name == get_active_profile_name():
            QMessageBox.warning(
                self, "Errore",
                "Non puoi eliminare il profilo attivo.\n"
                "Attiva un altro profilo prima di eliminare questo."
            )
            return
        r = QMessageBox.question(
            self, "Eliminare profilo?",
            f"Eliminare il profilo '{name}'?\n\n"
            "La configurazione viene rimossa.\n"
            "I dati nella cartella condivisa NON vengono eliminati.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if r == QMessageBox.StandardButton.Yes:
            delete_profile(name)
            self._refresh_list()

    def _copy_profile(self):
        src = self._selected_name()
        if not src:
            return
        dlg = CopyProfileDialog(src, self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        try:
            new_data = copy_profile(src, dlg.dst_name)
            # Sovrascrivi shared_root con il nuovo percorso
            new_data["shared_root"] = dlg.new_shared_root
            save_profile(dlg.dst_name, new_data)

            if dlg.copy_data:
                src_data = load_profile(src)
                src_root = src_data.get("shared_root", "")
                if src_root:
                    self._copy_shared_data(src_root, dlg.new_shared_root)

            self._refresh_list()
            QMessageBox.information(
                self, "OK", f"Profilo '{dlg.dst_name}' creato"
            )
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))

    def _copy_shared_data(self, src_root: str, dst_root: str):
        """Copia struttura cartella condivisa escludendo il database."""
        src = Path(src_root)
        dst = Path(dst_root)
        dst.mkdir(parents=True, exist_ok=True)

        for subdir in ("archive", "workspace", "thumbnails", "config", "temp"):
            src_dir = src / subdir
            dst_dir = dst / subdir
            if src_dir.exists():
                shutil.copytree(str(src_dir), str(dst_dir), dirs_exist_ok=True)
            else:
                dst_dir.mkdir(parents=True, exist_ok=True)
        # Crea directory database vuota (senza copiare il DB)
        (dst / "database").mkdir(parents=True, exist_ok=True)

    def _save_profile(self):
        old_name = self._selected_name()
        if not old_name:
            QMessageBox.warning(self, "Errore", "Seleziona un profilo dalla lista")
            return
        new_name = self.txt_name.text().strip()
        shared = self.txt_shared.text().strip()
        if not new_name:
            QMessageBox.warning(self, "Errore", "Inserire un nome per il profilo")
            return

        # Rinomina se il nome è cambiato
        if new_name != old_name:
            if new_name in get_profile_names():
                QMessageBox.warning(
                    self, "Errore", f"Il profilo '{new_name}' esiste già"
                )
                return
            rename_profile(old_name, new_name)

        # Aggiorna shared_root
        data = load_profile(new_name)
        data["shared_root"] = shared
        save_profile(new_name, data)

        self._refresh_list()
        QMessageBox.information(self, "OK", "Profilo salvato")

    def _activate_profile(self):
        name = self._selected_name()
        if not name:
            return
        data = load_profile(name)
        if not data.get("shared_root"):
            QMessageBox.warning(
                self, "Errore",
                "Configurare il percorso condiviso prima di attivare il profilo."
            )
            return
        if name == get_active_profile_name():
            QMessageBox.information(self, "Info", f"'{name}' è già il profilo attivo")
            return
        set_active_profile(name)
        self._refresh_list()
        self.profile_switched.emit(name)

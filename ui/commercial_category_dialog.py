# =============================================================================
#  ui/commercial_category_dialog.py  –  Gestione categorie/sottocategorie
# =============================================================================
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QListWidget, QListWidgetItem, QSplitter,
    QWidget, QTabWidget, QMessageBox, QFrame,
)
from PyQt6.QtCore import Qt

from ui.session import session


class _CategoryPanel(QWidget):
    """
    Pannello riutilizzabile per un tipo di articolo commerciale (5 o 6).
    Layout: QSplitter orizzontale
      - sinistra: lista categorie + form (codice auto, descrizione) + bottoni
      - destra:   lista sottocategorie + form (codice auto, descrizione, template) + bottoni
    """

    def __init__(self, item_type: str, parent=None):
        super().__init__(parent)
        self.item_type = item_type  # 'commerciale' | 'normalizzato'
        self._build_ui()
        self._load_categories()

    # ------------------------------------------------------------------
    # Costruzione UI
    # ------------------------------------------------------------------

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(6, 6, 6, 6)
        root.setSpacing(4)

        splitter = QSplitter(Qt.Orientation.Horizontal)

        # ---- Pannello categorie (sinistra) ----------------------------
        cat_widget = QWidget()
        cat_lay = QVBoxLayout(cat_widget)
        cat_lay.setContentsMargins(0, 0, 4, 0)
        cat_lay.setSpacing(4)

        lbl_cat = QLabel("CATEGORIE")
        lbl_cat.setStyleSheet("font-weight:bold; font-size:11px; color:#aaa;")
        cat_lay.addWidget(lbl_cat)

        self.lst_cat = QListWidget()
        self.lst_cat.currentItemChanged.connect(self._on_cat_selected)
        cat_lay.addWidget(self.lst_cat)

        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setFrameShadow(QFrame.Shadow.Sunken)
        cat_lay.addWidget(sep)

        form_cat = QVBoxLayout()
        form_cat.setSpacing(3)

        row_code = QHBoxLayout()
        row_code.addWidget(QLabel("Codice:"))
        self.txt_cat_code = QLineEdit()
        self.txt_cat_code.setPlaceholderText("es. VIT")
        self.txt_cat_code.setMaximumWidth(80)
        row_code.addWidget(self.txt_cat_code)
        row_code.addStretch()
        form_cat.addLayout(row_code)

        row_desc = QHBoxLayout()
        row_desc.addWidget(QLabel("Descrizione:"))
        self.txt_cat_desc = QLineEdit()
        self.txt_cat_desc.setPlaceholderText("Nome categoria…")
        row_desc.addWidget(self.txt_cat_desc)
        form_cat.addLayout(row_desc)

        cat_lay.addLayout(form_cat)

        btn_row_cat = QHBoxLayout()
        self.btn_cat_new = QPushButton("+ Nuova")
        self.btn_cat_new.setFixedHeight(26)
        self.btn_cat_new.clicked.connect(self._create_category)
        self.btn_cat_edit = QPushButton("Modifica")
        self.btn_cat_edit.setFixedHeight(26)
        self.btn_cat_edit.clicked.connect(self._edit_category)
        self.btn_cat_del = QPushButton("Elimina")
        self.btn_cat_del.setFixedHeight(26)
        self.btn_cat_del.clicked.connect(self._delete_category)
        btn_row_cat.addWidget(self.btn_cat_new)
        btn_row_cat.addWidget(self.btn_cat_edit)
        btn_row_cat.addWidget(self.btn_cat_del)
        cat_lay.addLayout(btn_row_cat)

        splitter.addWidget(cat_widget)

        # ---- Pannello sottocategorie (destra) -------------------------
        sub_widget = QWidget()
        sub_lay = QVBoxLayout(sub_widget)
        sub_lay.setContentsMargins(4, 0, 0, 0)
        sub_lay.setSpacing(4)

        lbl_sub = QLabel("SOTTOCATEGORIE")
        lbl_sub.setStyleSheet("font-weight:bold; font-size:11px; color:#aaa;")
        sub_lay.addWidget(lbl_sub)

        self.lst_sub = QListWidget()
        self.lst_sub.currentItemChanged.connect(self._on_sub_selected)
        sub_lay.addWidget(self.lst_sub)

        sep2 = QFrame()
        sep2.setFrameShape(QFrame.Shape.HLine)
        sep2.setFrameShadow(QFrame.Shadow.Sunken)
        sub_lay.addWidget(sep2)

        form_sub = QVBoxLayout()
        form_sub.setSpacing(3)

        row_scode = QHBoxLayout()
        row_scode.addWidget(QLabel("Codice:"))
        self.txt_sub_code = QLineEdit()
        self.txt_sub_code.setPlaceholderText("es. ISO")
        self.txt_sub_code.setMaximumWidth(80)
        row_scode.addWidget(self.txt_sub_code)
        row_scode.addStretch()
        form_sub.addLayout(row_scode)

        row_sdesc = QHBoxLayout()
        row_sdesc.addWidget(QLabel("Descrizione:"))
        self.txt_sub_desc = QLineEdit()
        self.txt_sub_desc.setPlaceholderText("Nome sottocategoria…")
        row_sdesc.addWidget(self.txt_sub_desc)
        form_sub.addLayout(row_sdesc)

        row_tpl = QHBoxLayout()
        row_tpl.addWidget(QLabel("Template:"))
        self.txt_sub_tpl = QLineEdit()
        self.txt_sub_tpl.setPlaceholderText("es. Vite M{size} x {length}")
        row_tpl.addWidget(self.txt_sub_tpl)
        form_sub.addLayout(row_tpl)

        sub_lay.addLayout(form_sub)

        btn_row_sub = QHBoxLayout()
        self.btn_sub_new = QPushButton("+ Nuova")
        self.btn_sub_new.setFixedHeight(26)
        self.btn_sub_new.clicked.connect(self._create_subcategory)
        self.btn_sub_edit = QPushButton("Modifica")
        self.btn_sub_edit.setFixedHeight(26)
        self.btn_sub_edit.clicked.connect(self._edit_subcategory)
        self.btn_sub_del = QPushButton("Elimina")
        self.btn_sub_del.setFixedHeight(26)
        self.btn_sub_del.clicked.connect(self._delete_subcategory)
        btn_row_sub.addWidget(self.btn_sub_new)
        btn_row_sub.addWidget(self.btn_sub_edit)
        btn_row_sub.addWidget(self.btn_sub_del)
        sub_lay.addLayout(btn_row_sub)

        splitter.addWidget(sub_widget)
        splitter.setSizes([320, 380])

        root.addWidget(splitter)

    # ------------------------------------------------------------------
    # Caricamento dati
    # ------------------------------------------------------------------

    def _load_categories(self):
        self.lst_cat.blockSignals(True)
        self.lst_cat.clear()
        cats = session.commercial.get_categories(item_type=self.item_type)
        for cat in cats:
            item = QListWidgetItem(f"{cat['code']}  {cat['description']}")
            item.setData(Qt.ItemDataRole.UserRole, cat["id"])
            self.lst_cat.addItem(item)
        self.lst_cat.blockSignals(False)
        self._update_cat_code_hint()
        self.lst_sub.clear()
        self._update_sub_code_hint(None)

    def _load_subcategories(self, cat_id: int):
        self.lst_sub.blockSignals(True)
        self.lst_sub.clear()
        subs = session.commercial.get_subcategories(cat_id)
        for sub in subs:
            item = QListWidgetItem(f"{sub['code']}  {sub['description']}")
            item.setData(Qt.ItemDataRole.UserRole, sub["id"])
            self.lst_sub.addItem(item)
        self.lst_sub.blockSignals(False)
        self._update_sub_code_hint(cat_id)

    # ------------------------------------------------------------------
    # Suggerimento codice (pre-compilazione modificabile)
    # ------------------------------------------------------------------

    def _update_cat_code_hint(self):
        """Pre-compila il campo codice con il suggerimento (solo se vuoto)."""
        if self.txt_cat_code.text().strip():
            return
        try:
            code = session.commercial.get_next_category_code(self.item_type)
            self.txt_cat_code.setText(code)
        except Exception:
            pass

    def _update_sub_code_hint(self, cat_id):
        """Pre-compila il campo codice sottocategoria (solo se vuoto)."""
        if cat_id is None:
            self.txt_sub_code.clear()
            return
        if self.txt_sub_code.text().strip():
            return
        try:
            code = session.commercial.get_next_subcategory_code(cat_id)
            self.txt_sub_code.setText(code)
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Selezione
    # ------------------------------------------------------------------

    def _on_cat_selected(self, current, _previous):
        if current is None:
            self.txt_cat_desc.clear()
            self.lst_sub.clear()
            self._update_sub_code_hint(None)
            return
        cat_id = current.data(Qt.ItemDataRole.UserRole)
        cat = session.commercial.get_category(cat_id)
        if cat:
            self.txt_cat_desc.setText(cat["description"])
        self._load_subcategories(cat_id)

    def _on_sub_selected(self, current, _previous):
        if current is None:
            self.txt_sub_desc.clear()
            self.txt_sub_tpl.clear()
            return
        sub_id = current.data(Qt.ItemDataRole.UserRole)
        sub = session.commercial.get_subcategory(sub_id)
        if sub:
            self.txt_sub_desc.setText(sub["description"])
            self.txt_sub_tpl.setText(sub.get("desc_template") or "")

    # ------------------------------------------------------------------
    # Azioni categorie
    # ------------------------------------------------------------------

    def _create_category(self):
        code = self.txt_cat_code.text().strip()
        desc = self.txt_cat_desc.text().strip()
        if not code:
            QMessageBox.warning(self, "Attenzione", "Inserire un codice categoria.")
            return
        if len(code) > 10:
            QMessageBox.warning(self, "Attenzione", "Il codice non può superare 10 caratteri.")
            return
        if not desc:
            QMessageBox.warning(self, "Attenzione", "Inserire una descrizione.")
            return
        try:
            session.commercial.create_category(code, desc, self.item_type)
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))
            return
        self.txt_cat_code.clear()
        self.txt_cat_desc.clear()
        self._load_categories()

    def _edit_category(self):
        item = self.lst_cat.currentItem()
        if not item:
            QMessageBox.information(self, "Seleziona", "Selezionare una categoria.")
            return
        desc = self.txt_cat_desc.text().strip()
        if not desc:
            QMessageBox.warning(self, "Attenzione", "Inserire una descrizione.")
            return
        cat_id = item.data(Qt.ItemDataRole.UserRole)
        try:
            session.commercial.update_category(cat_id, desc)
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))
            return
        self._load_categories()

    def _delete_category(self):
        item = self.lst_cat.currentItem()
        if not item:
            return
        cat_id = item.data(Qt.ItemDataRole.UserRole)
        cat = session.commercial.get_category(cat_id)
        name = cat["description"] if cat else "?"
        reply = QMessageBox.question(
            self, "Conferma eliminazione",
            f"Eliminare la categoria «{name}» e tutte le sue sottocategorie?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        try:
            session.commercial.delete_category(cat_id)
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))
            return
        self._load_categories()

    # ------------------------------------------------------------------
    # Azioni sottocategorie
    # ------------------------------------------------------------------

    def _current_cat_id(self):
        item = self.lst_cat.currentItem()
        if not item:
            return None
        return item.data(Qt.ItemDataRole.UserRole)

    def _create_subcategory(self):
        cat_id = self._current_cat_id()
        if cat_id is None:
            QMessageBox.information(self, "Seleziona", "Selezionare prima una categoria.")
            return
        code = self.txt_sub_code.text().strip()
        desc = self.txt_sub_desc.text().strip()
        if not code:
            QMessageBox.warning(self, "Attenzione", "Inserire un codice sottocategoria.")
            return
        if len(code) > 10:
            QMessageBox.warning(self, "Attenzione", "Il codice non può superare 10 caratteri.")
            return
        if not desc:
            QMessageBox.warning(self, "Attenzione", "Inserire una descrizione.")
            return
        tpl = self.txt_sub_tpl.text().strip()
        try:
            session.commercial.create_subcategory(cat_id, code, desc, tpl)
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))
            return
        self.txt_sub_code.clear()
        self.txt_sub_desc.clear()
        self.txt_sub_tpl.clear()
        self._load_subcategories(cat_id)

    def _edit_subcategory(self):
        item = self.lst_sub.currentItem()
        if not item:
            QMessageBox.information(self, "Seleziona", "Selezionare una sottocategoria.")
            return
        desc = self.txt_sub_desc.text().strip()
        if not desc:
            QMessageBox.warning(self, "Attenzione", "Inserire una descrizione.")
            return
        tpl = self.txt_sub_tpl.text().strip()
        sub_id = item.data(Qt.ItemDataRole.UserRole)
        try:
            session.commercial.update_subcategory(sub_id, desc, tpl)
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))
            return
        cat_id = self._current_cat_id()
        if cat_id:
            self._load_subcategories(cat_id)

    def _delete_subcategory(self):
        item = self.lst_sub.currentItem()
        if not item:
            return
        sub_id = item.data(Qt.ItemDataRole.UserRole)
        sub = session.commercial.get_subcategory(sub_id)
        name = sub["description"] if sub else "?"
        reply = QMessageBox.question(
            self, "Conferma eliminazione",
            f"Eliminare la sottocategoria «{name}»?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        try:
            session.commercial.delete_subcategory(sub_id)
        except Exception as e:
            QMessageBox.critical(self, "Errore", str(e))
            return
        cat_id = self._current_cat_id()
        if cat_id:
            self._load_subcategories(cat_id)


# =============================================================================
#  Dialog principale
# =============================================================================

class CommercialCategoryDialog(QDialog):
    """Dialog con due tab (Commerciali / Normalizzati), ognuno con due box affiancati."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Gestione Categorie Commerciali")
        self.setMinimumSize(780, 500)
        self._build_ui()

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(10, 10, 10, 10)

        tabs = QTabWidget()
        self._panel_com = _CategoryPanel("commerciale")
        self._panel_nor = _CategoryPanel("normalizzato")
        tabs.addTab(self._panel_com, "Commerciali (5)")
        tabs.addTab(self._panel_nor, "Normalizzati (6)")
        lay.addWidget(tabs)

        btn_close = QPushButton("Chiudi")
        btn_close.setFixedWidth(100)
        btn_close.clicked.connect(self.accept)
        row = QHBoxLayout()
        row.addStretch()
        row.addWidget(btn_close)
        lay.addLayout(row)

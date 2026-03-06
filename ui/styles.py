# =============================================================================
#  ui/styles.py  –  Stili e tema applicazione
# =============================================================================

DARK_THEME = """
QWidget {
    background-color: #1e1e2e;
    color: #cdd6f4;
    font-family: 'Segoe UI', Arial, sans-serif;
    font-size: 13px;
}

QMainWindow, QDialog {
    background-color: #1e1e2e;
}

/* ---- MenuBar ---- */
QMenuBar {
    background-color: #181825;
    color: #cdd6f4;
    border-bottom: 1px solid #313244;
}
QMenuBar::item:selected { background-color: #313244; }
QMenu {
    background-color: #181825;
    border: 1px solid #45475a;
}
QMenu::item:selected { background-color: #313244; }

/* ---- ToolBar ---- */
QToolBar {
    background-color: #181825;
    border-bottom: 1px solid #313244;
    spacing: 4px;
    padding: 2px;
}
QToolButton {
    background-color: transparent;
    border: 1px solid transparent;
    border-radius: 4px;
    padding: 4px 8px;
    color: #cdd6f4;
}
QToolButton:hover { background-color: #313244; border-color: #45475a; }
QToolButton:pressed { background-color: #45475a; }

/* ---- Buttons ---- */
QPushButton {
    background-color: #313244;
    color: #cdd6f4;
    border: 1px solid #45475a;
    border-radius: 5px;
    padding: 6px 16px;
    min-width: 80px;
}
QPushButton:hover { background-color: #45475a; border-color: #89b4fa; }
QPushButton:pressed { background-color: #585b70; }
QPushButton:disabled { color: #585b70; border-color: #313244; }

QPushButton#btn_primary {
    background-color: #89b4fa;
    color: #1e1e2e;
    font-weight: bold;
}
QPushButton#btn_primary:hover { background-color: #b4d0ff; }

QPushButton#btn_success {
    background-color: #a6e3a1;
    color: #1e1e2e;
    font-weight: bold;
}
QPushButton#btn_danger {
    background-color: #f38ba8;
    color: #1e1e2e;
    font-weight: bold;
}
QPushButton#btn_warning {
    background-color: #fab387;
    color: #1e1e2e;
    font-weight: bold;
}

/* ---- LineEdit, ComboBox, TextEdit ---- */
QLineEdit, QTextEdit, QPlainTextEdit, QSpinBox, QDoubleSpinBox {
    background-color: #313244;
    border: 1px solid #45475a;
    border-radius: 4px;
    padding: 4px 8px;
    color: #cdd6f4;
    selection-background-color: #89b4fa;
    selection-color: #1e1e2e;
}
QLineEdit:focus, QTextEdit:focus, QPlainTextEdit:focus {
    border-color: #89b4fa;
}

QComboBox {
    background-color: #313244;
    border: 1px solid #45475a;
    border-radius: 4px;
    padding: 4px 8px;
    color: #cdd6f4;
}
QComboBox:focus { border-color: #89b4fa; }
QComboBox::drop-down { border: none; width: 20px; }
QComboBox QAbstractItemView {
    background-color: #313244;
    border: 1px solid #45475a;
    selection-background-color: #45475a;
}

/* ---- Tables ---- */
QTableWidget, QTableView {
    background-color: #181825;
    alternate-background-color: #1e1e2e;
    border: 1px solid #313244;
    gridline-color: #313244;
    selection-background-color: #313244;
    selection-color: #cdd6f4;
}
QHeaderView::section {
    background-color: #181825;
    color: #89b4fa;
    font-weight: bold;
    padding: 6px;
    border: none;
    border-right: 1px solid #313244;
    border-bottom: 1px solid #45475a;
}
QTableWidget::item, QTableView::item { padding: 4px 8px; }
QTableWidget::item:selected, QTableView::item:selected {
    background-color: #313244;
    color: #cdd6f4;
}

/* ---- TreeWidget ---- */
QTreeWidget {
    background-color: #181825;
    border: 1px solid #313244;
    alternate-background-color: #1e1e2e;
}
QTreeWidget::item { padding: 3px; }
QTreeWidget::item:selected { background-color: #313244; }
QTreeWidget::item:hover { background-color: #232334; }

/* ---- Tab ---- */
QTabWidget::pane {
    border: 1px solid #313244;
    border-radius: 4px;
    background-color: #1e1e2e;
}
QTabBar::tab {
    background-color: #181825;
    border: 1px solid #313244;
    border-bottom: none;
    border-radius: 4px 4px 0 0;
    padding: 6px 14px;
    color: #a6adc8;
}
QTabBar::tab:selected {
    background-color: #313244;
    color: #cdd6f4;
    font-weight: bold;
}
QTabBar::tab:hover { background-color: #232334; }

/* ---- StatusBar ---- */
QStatusBar {
    background-color: #181825;
    border-top: 1px solid #313244;
    color: #a6adc8;
}

/* ---- GroupBox ---- */
QGroupBox {
    border: 1px solid #45475a;
    border-radius: 6px;
    margin-top: 10px;
    padding-top: 10px;
    color: #89b4fa;
    font-weight: bold;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 4px;
}

/* ---- Scrollbar ---- */
QScrollBar:vertical {
    background: #181825;
    width: 10px;
    border-radius: 5px;
}
QScrollBar::handle:vertical {
    background: #45475a;
    border-radius: 5px;
    min-height: 20px;
}
QScrollBar::handle:vertical:hover { background: #585b70; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }

QScrollBar:horizontal {
    background: #181825;
    height: 10px;
    border-radius: 5px;
}
QScrollBar::handle:horizontal {
    background: #45475a;
    border-radius: 5px;
    min-width: 20px;
}
QScrollBar::handle:horizontal:hover { background: #585b70; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0; }

/* ---- Splitter ---- */
QSplitter::handle { background-color: #313244; }
QSplitter::handle:horizontal { width: 2px; }
QSplitter::handle:vertical { height: 2px; }

/* ---- Label ---- */
QLabel#title_label {
    font-size: 16px;
    font-weight: bold;
    color: #89b4fa;
}
QLabel#subtitle_label {
    font-size: 11px;
    color: #a6adc8;
}
"""


STATE_BADGE_STYLE = {
    "In Lavorazione": "background:#2196F3;color:white;border-radius:4px;padding:2px 8px;",
    "In Revisione":   "background:#FF9800;color:white;border-radius:4px;padding:2px 8px;",
    "Rilasciato":     "background:#4CAF50;color:white;border-radius:4px;padding:2px 8px;",
    "Obsoleto":       "background:#757575;color:white;border-radius:4px;padding:2px 8px;",
}

TYPE_ICON = {
    "Parte":   "🔩",
    "Assieme": "⚙️",
    "Disegno": "📐",
}

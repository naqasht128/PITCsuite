import os
from PySide6.QtGui import QPalette, QColor, QFont
from PySide6.QtWidgets import QLabel, QFrame, QMessageBox, QPushButton
from PySide6.QtCore import Qt

def configure_app(app):
    app.setStyle("Fusion")
    
    # Modern readable system font
    font = app.font()
    font.setFamily("Segoe UI")
    font.setPointSize(9)
    app.setFont(font)
    
    # High-contrast Dark Theme Palette
    palette = QPalette()
    colors = {
        "Window": "#0d1117",
        "WindowText": "#f0f6fc",
        "Base": "#090d14",
        "AlternateBase": "#131822",
        "ToolTipBase": "#1a2233",
        "ToolTipText": "#f8fafc",
        "Text": "#f0f6fc",
        "Button": "#161c28",
        "ButtonText": "#f0f6fc",
        "BrightText": "#ffffff",
        "Highlight": "#2563eb",
        "HighlightedText": "#ffffff",
        "PlaceholderText": "#64748b",
        "Link": "#38bdf8",
    }
    for role, color in colors.items():
        palette.setColor(getattr(QPalette.ColorRole, role), QColor(color))
        
    for role in (QPalette.ColorRole.Text, QPalette.ColorRole.WindowText, QPalette.ColorRole.ButtonText):
        palette.setColor(QPalette.ColorGroup.Disabled, role, QColor("#475569"))
        
    app.setPalette(palette)


DARK_STYLE = """
/* Global Application Stylesheet */
QWidget {
    color: #f0f6fc;
    font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, 'Inter', sans-serif;
    font-size: 9pt;
}

QMainWindow, QDialog {
    background-color: #0d1117;
}

/* Scroll Areas */
QScrollArea {
    background: transparent;
    border: none;
}

QScrollArea > QWidget > QWidget {
    background: transparent;
}

/* Headings & Dividers */
QFrame[frameShape="4"], QFrame[frameShape="5"] {
    border: none;
    background-color: #1e2638;
    height: 1px;
    max-height: 1px;
}

/* Card-like Group Boxes */
QGroupBox {
    background-color: #121722;
    border: 1px solid #1e2638;
    border-radius: 8px;
    margin-top: 18px;
    padding: 14px 12px 12px 12px;
    font-weight: 600;
}

QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 12px;
    padding: 2px 8px;
    color: #60a5fa;
    background-color: #121722;
    border: 1px solid #1e2638;
    border-radius: 4px;
    font-size: 8.5pt;
    font-weight: 600;
}

/* Text Inputs, Spinners, Dates */
QLineEdit, QSpinBox, QDoubleSpinBox, QDateEdit {
    background-color: #0b0e15;
    color: #f0f6fc;
    border: 1px solid #222c3d;
    border-radius: 6px;
    padding: 6px 10px;
    font-size: 9pt;
    min-height: 20px;
    selection-background-color: #2563eb;
    selection-color: #ffffff;
}

QLineEdit:hover, QSpinBox:hover, QDoubleSpinBox:hover, QDateEdit:hover {
    border-color: #334155;
}

QLineEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus, QDateEdit:focus {
    border: 1px solid #3b82f6;
    background-color: #0e131d;
}

QLineEdit:read-only {
    background-color: #131720;
    color: #94a3b8;
    border-color: #1e2638;
}

QLineEdit:disabled, QSpinBox:disabled, QDoubleSpinBox:disabled, QDateEdit:disabled {
    background-color: #10141d;
    color: #475569;
    border-color: #1a2232;
}

/* SpinBox controls */
QSpinBox::up-button, QDoubleSpinBox::up-button, QDateEdit::up-button {
    subcontrol-origin: border;
    subcontrol-position: top right;
    width: 18px;
    border-left: 1px solid #222c3d;
    border-bottom: 1px solid #222c3d;
    border-top-right-radius: 5px;
    background: #141a24;
}

QSpinBox::up-button:hover, QDoubleSpinBox::up-button:hover, QDateEdit::up-button:hover {
    background: #2563eb;
}

QSpinBox::down-button, QDoubleSpinBox::down-button, QDateEdit::down-button {
    subcontrol-origin: border;
    subcontrol-position: bottom right;
    width: 18px;
    border-left: 1px solid #222c3d;
    border-bottom-right-radius: 5px;
    background: #141a24;
}

QSpinBox::down-button:hover, QDoubleSpinBox::down-button:hover, QDateEdit::down-button:hover {
    background: #2563eb;
}

/* Combo Boxes */
QComboBox {
    background-color: #0b0e15;
    color: #f0f6fc;
    border: 1px solid #222c3d;
    border-radius: 6px;
    padding: 6px 10px;
    font-size: 9pt;
    min-height: 20px;
}

QComboBox:hover {
    border-color: #334155;
}

QComboBox:focus {
    border: 1px solid #3b82f6;
    background-color: #0e131d;
}

QComboBox::drop-down {
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 24px;
    border-left: 1px solid #222c3d;
    border-top-right-radius: 5px;
    border-bottom-right-radius: 5px;
    background: #141a24;
}

QComboBox::drop-down:hover {
    background: #1e293b;
}

QComboBox QAbstractItemView {
    background-color: #121722;
    color: #f0f6fc;
    border: 1px solid #2a3449;
    selection-background-color: #1e293b;
    selection-color: #38bdf8;
    padding: 4px;
    border-radius: 6px;
    outline: none;
}

/* Checkboxes and Radio Buttons */
QCheckBox, QRadioButton {
    color: #cbd5e1;
    spacing: 8px;
    font-size: 9pt;
}

QCheckBox:hover, QRadioButton:hover {
    color: #f8fafc;
}

QCheckBox::indicator, QRadioButton::indicator {
    width: 16px;
    height: 16px;
    border: 1px solid #2c384c;
    border-radius: 4px;
    background: #0b0e15;
}

QRadioButton::indicator {
    border-radius: 8px;
}

QCheckBox::indicator:hover, QRadioButton::indicator:hover {
    border-color: #3b82f6;
    background: #0f1420;
}

QCheckBox::indicator:checked {
    background-color: #2563eb;
    border-color: #3b82f6;
}

QRadioButton::indicator:checked {
    background-color: #2563eb;
    border: 4px solid #0b0e15;
}

/* Progress Bars */
QProgressBar {
    background-color: #0b0e15;
    border: 1px solid #1e2638;
    border-radius: 6px;
    text-align: center;
    color: #ffffff;
    font-size: 8.5pt;
    font-weight: 600;
    min-height: 14px;
    max-height: 14px;
}

QProgressBar::chunk {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #2563eb, stop:1 #38bdf8);
    border-radius: 5px;
}

/* Terminal & Log Text Areas */
QTextEdit, QPlainTextEdit {
    background-color: #080b11;
    color: #38bdf8;
    border: 1px solid #1c2333;
    border-radius: 6px;
    padding: 8px;
    font-family: 'Cascadia Code', 'Consolas', monospace;
    font-size: 8.5pt;
    selection-background-color: #1e3a8a;
    selection-color: #ffffff;
}

/* Lists and Tables */
QListWidget, QTableWidget, QTableView, QTreeView {
    background-color: #0b0e15;
    color: #e2e8f0;
    border: 1px solid #1e2638;
    border-radius: 6px;
    gridline-color: #1a2232;
    selection-background-color: #1e293b;
    selection-color: #60a5fa;
    outline: none;
}

QListWidget::item, QTableWidget::item {
    padding: 5px 8px;
    border-radius: 4px;
}

QListWidget::item:hover, QTableWidget::item:hover {
    background-color: #141b27;
}

QListWidget::item:selected, QTableWidget::item:selected {
    background-color: #1e293b;
    color: #60a5fa;
}

QHeaderView::section {
    background-color: #121722;
    color: #94a3b8;
    border: none;
    border-bottom: 1px solid #1e2638;
    border-right: 1px solid #1e2638;
    padding: 6px 8px;
    font-weight: 600;
    font-size: 8.5pt;
}

QTableCornerButton::section {
    background-color: #121722;
    border: none;
}

/* Scrollbars */
QScrollBar:vertical {
    background: transparent;
    width: 8px;
    margin: 0;
}

QScrollBar::handle:vertical {
    background: #232d3f;
    min-height: 24px;
    border-radius: 4px;
}

QScrollBar::handle:vertical:hover {
    background: #3b4961;
}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical,
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
    height: 0px;
    background: none;
}

QScrollBar:horizontal {
    background: transparent;
    height: 8px;
    margin: 0;
}

QScrollBar::handle:horizontal {
    background: #232d3f;
    min-width: 24px;
    border-radius: 4px;
}

QScrollBar::handle:horizontal:hover {
    background: #3b4961;
}

QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal,
QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
    width: 0px;
    background: none;
}

/* Tabs */
QTabWidget::pane {
    border: 1px solid #1e2638;
    background-color: #121722;
    border-radius: 6px;
}

QTabBar::tab {
    background: #141924;
    color: #94a3b8;
    padding: 7px 16px;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
    margin-right: 2px;
    border: 1px solid #1e2638;
    border-bottom: none;
    font-weight: 500;
}

QTabBar::tab:selected {
    background: #121722;
    color: #60a5fa;
    font-weight: 600;
    border-bottom: 2px solid #3b82f6;
}

QTabBar::tab:hover:!selected {
    background: #1a2232;
    color: #f1f5f9;
}

/* Tooltips */
QToolTip {
    background-color: #1a2233;
    color: #f8fafc;
    border: 1px solid #2e3a52;
    border-radius: 5px;
    padding: 5px 8px;
    font-size: 8.5pt;
}

/* Default PushButton */
QPushButton {
    background: #161c28;
    color: #e2e8f0;
    border: 1px solid #253043;
    border-radius: 6px;
    padding: 6px 14px;
    font-size: 9pt;
    font-weight: 500;
}

QPushButton:hover {
    background: #1f2738;
    border-color: #3b82f6;
    color: #ffffff;
}

QPushButton:pressed {
    background: #121721;
}

QPushButton:disabled {
    background: #141822;
    color: #475569;
    border-color: #1e2638;
}

/* Primary role button */
QPushButton[role=primary] {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #10b981, stop:1 #059669);
    color: #ffffff;
    border: 1px solid #34d399;
    border-radius: 6px;
    padding: 7px 18px;
    font-weight: 600;
    font-size: 9pt;
}

QPushButton[role=primary]:hover {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #34d399, stop:1 #10b981);
    border-color: #6ee7b7;
}

QPushButton[role=primary]:pressed {
    background: #047857;
    border-color: #059669;
}

/* Secondary role button */
QPushButton[role=secondary] {
    background: #161c28;
    color: #e2e8f0;
    border: 1px solid #253043;
    border-radius: 6px;
    padding: 6px 14px;
    font-weight: 500;
}

QPushButton[role=secondary]:hover {
    background: #1f2738;
    border-color: #3b82f6;
    color: #ffffff;
}
"""

START_BTN = """
QPushButton {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #10b981, stop:1 #059669);
    color: #ffffff;
    border: 1px solid #34d399;
    border-radius: 6px;
    padding: 7px 18px;
    font-weight: 600;
    font-size: 9pt;
}
QPushButton:hover {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #34d399, stop:1 #10b981);
    border-color: #6ee7b7;
}
QPushButton:pressed {
    background: #047857;
    border-color: #059669;
}
QPushButton:disabled {
    background: #141822;
    color: #475569;
    border: 1px solid #1e2638;
}
"""

STOP_BTN = """
QPushButton {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #ef4444, stop:1 #dc2626);
    color: #ffffff;
    border: 1px solid #f87171;
    border-radius: 6px;
    padding: 7px 16px;
    font-weight: 600;
    font-size: 9pt;
}
QPushButton:hover {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #f87171, stop:1 #ef4444);
    border-color: #fca5a5;
}
QPushButton:pressed {
    background: #b91c1c;
    border-color: #991b1b;
}
QPushButton:disabled {
    background: #141822;
    color: #475569;
    border: 1px solid #1e2638;
}
"""

NAV_STYLE = """
QPushButton {
    background: transparent;
    color: #94a3b8;
    border: none;
    border-left: 3px solid transparent;
    border-radius: 6px;
    padding: 8px 10px;
    text-align: left;
    font-size: 8.5pt;
    font-weight: 500;
    margin: 1px 6px;
}
QPushButton:hover {
    background: #18202d;
    color: #f1f5f9;
}
QPushButton[active="true"] {
    background: #1e293b;
    color: #38bdf8;
    border-left-color: #38bdf8;
    font-weight: 600;
}
"""


def lbl(text, bold=False, color=None):
    w = QLabel(text)
    w.setWordWrap(True)
    if bold:
        f = w.font()
        f.setBold(True)
        w.setFont(f)
    if color:
        w.setStyleSheet(f"color:{color};")
    return w


def hline():
    f = QFrame()
    f.setFrameShape(QFrame.HLine)
    f.setStyleSheet("background-color: #1e2638; border: none; height: 1px; max-height: 1px;")
    return f


def notify(parent, title, message, critical=False, file_path=None, folder_path=None):
    """Show a result without blocking the Qt event loop with exec()."""
    box = QMessageBox(QMessageBox.Critical if critical else QMessageBox.Information,
                      title, message, parent=parent)
    box.setAttribute(Qt.WA_DeleteOnClose)
    box.setStyleSheet("""
        QMessageBox {
            background-color: #121722;
            color: #f0f6fc;
        }
        QLabel {
            color: #f0f6fc;
            font-size: 9pt;
        }
        QPushButton {
            background: #1e293b;
            color: #e2e8f0;
            border: 1px solid #334155;
            border-radius: 6px;
            padding: 6px 16px;
            font-weight: 500;
        }
        QPushButton:hover {
            background: #2563eb;
            color: #ffffff;
            border-color: #3b82f6;
        }
    """)
    
    # Add an OK button as the default
    box.addButton(QMessageBox.Ok)
    
    if file_path:
        open_file_btn = box.addButton("Open the file", QMessageBox.ActionRole)
        open_file_btn.clicked.connect(lambda _=False, p=file_path: os.startfile(p) if os.path.exists(p) else None)
    
    if folder_path:
        open_folder_btn = box.addButton("Open the Folder", QMessageBox.ActionRole)
        open_folder_btn.clicked.connect(lambda _=False, p=folder_path: os.startfile(p) if os.path.exists(p) else None)
        
    active = getattr(parent, "_active_notifications", None)
    if active is None:
        active = parent._active_notifications = []
    active.append(box)
    box.finished.connect(lambda: active.remove(box) if box in active else None)
    box.open()
    return box

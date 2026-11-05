import os
from PySide6.QtGui import QPalette, QColor, QFont
from PySide6.QtWidgets import QLabel, QFrame, QMessageBox, QPushButton
from PySide6.QtCore import Qt

def configure_app(app):
    app.setStyle("Fusion")
    
    # Antigravity exact system font
    font = app.font()
    font.setFamily("Segoe UI")
    font.setPointSize(9)
    app.setFont(font)
    
    # Antigravity Dark+ Palette
    palette = QPalette()
    colors = {
        "Window": "#1e1e1e",
        "WindowText": "#cccccc",
        "Base": "#181818",
        "AlternateBase": "#252526",
        "ToolTipBase": "#252526",
        "ToolTipText": "#cccccc",
        "Text": "#cccccc",
        "Button": "#2d2d2d",
        "ButtonText": "#cccccc",
        "BrightText": "#ffffff",
        "Highlight": "#0078d4",
        "HighlightedText": "#ffffff",
        "PlaceholderText": "#808080",
        "Link": "#3794ff",
    }
    for role, color in colors.items():
        palette.setColor(getattr(QPalette.ColorRole, role), QColor(color))
        
    for role in (QPalette.ColorRole.Text, QPalette.ColorRole.WindowText, QPalette.ColorRole.ButtonText):
        palette.setColor(QPalette.ColorGroup.Disabled, role, QColor("#656565"))
        
    app.setPalette(palette)


DARK_STYLE = """
/* Antigravity Exact Match Dark Theme */
QWidget {
    color: #cccccc;
    font-family: 'Segoe UI', 'Google Sans', -apple-system, BlinkMacSystemFont, sans-serif;
    font-size: 8.5pt;
}

QMainWindow, QDialog {
    background-color: #1e1e1e;
}

/* Scroll Areas */
QScrollArea {
    background: transparent;
    border: none;
}

QScrollArea > QWidget > QWidget {
    background: transparent;
}

/* Dividers */
QFrame[frameShape="4"], QFrame[frameShape="5"] {
    border: none;
    background-color: #2b2b2b;
    height: 1px;
    max-height: 1px;
}

/* Antigravity Containers / GroupBoxes */
QGroupBox {
    background-color: #252526;
    border: 1px solid #333333;
    border-radius: 6px;
    margin-top: 13px;
    padding: 11px 9px 9px 9px;
    font-weight: 600;
}

QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 10px;
    padding: 2px 8px;
    color: #e3e3e3;
    background-color: #252526;
    border: 1px solid #3c3c3c;
    border-radius: 4px;
    font-size: 8pt;
    font-weight: 600;
}

/* Inputs, Spinners, Dates */
QLineEdit, QSpinBox, QDoubleSpinBox, QDateEdit {
    background-color: #313131;
    color: #cccccc;
    border: 1px solid #3c3c3c;
    border-radius: 4px;
    padding: 5px 8px;
    font-size: 8.5pt;
    min-height: 19px;
    selection-background-color: #264f78;
    selection-color: #ffffff;
}

QLineEdit:hover, QSpinBox:hover, QDoubleSpinBox:hover, QDateEdit:hover {
    border-color: #4a4a4a;
}

QLineEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus, QDateEdit:focus {
    border: 1px solid #0078d4;
    background-color: #313131;
    color: #ffffff;
}

QLineEdit:read-only {
    background-color: #2b2b2b;
    color: #9d9d9d;
    border-color: #383838;
}

QLineEdit:disabled, QSpinBox:disabled, QDoubleSpinBox:disabled, QDateEdit:disabled {
    background-color: #252526;
    color: #656565;
    border-color: #2f2f2f;
}

/* SpinBox controls */
QSpinBox::up-button, QDoubleSpinBox::up-button, QDateEdit::up-button {
    subcontrol-origin: border;
    subcontrol-position: top right;
    width: 16px;
    border-left: 1px solid #3c3c3c;
    border-bottom: 1px solid #3c3c3c;
    border-top-right-radius: 3px;
    background: #2d2d2d;
}

QSpinBox::up-button:hover, QDoubleSpinBox::up-button:hover, QDateEdit::up-button:hover {
    background: #0078d4;
}

QSpinBox::down-button, QDoubleSpinBox::down-button, QDateEdit::down-button {
    subcontrol-origin: border;
    subcontrol-position: bottom right;
    width: 16px;
    border-left: 1px solid #3c3c3c;
    border-bottom-right-radius: 3px;
    background: #2d2d2d;
}

QSpinBox::down-button:hover, QDoubleSpinBox::down-button:hover, QDateEdit::down-button:hover {
    background: #0078d4;
}

/* Combo Boxes */
QComboBox {
    background-color: #313131;
    color: #cccccc;
    border: 1px solid #3c3c3c;
    border-radius: 4px;
    padding: 5px 8px;
    font-size: 8.5pt;
    min-height: 19px;
}

QComboBox:hover {
    border-color: #4a4a4a;
}

QComboBox:focus {
    border: 1px solid #0078d4;
    background-color: #313131;
    color: #ffffff;
}

QComboBox::drop-down {
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 22px;
    border-left: 1px solid #3c3c3c;
    border-top-right-radius: 3px;
    border-bottom-right-radius: 3px;
    background: #2d2d2d;
}

QComboBox::drop-down:hover {
    background: #383838;
}

QComboBox QAbstractItemView {
    background-color: #252526;
    color: #cccccc;
    border: 1px solid #3c3c3c;
    selection-background-color: #04395e;
    selection-color: #ffffff;
    padding: 3px;
    border-radius: 4px;
    outline: none;
}

/* Checkboxes and Radio Buttons */
QCheckBox, QRadioButton {
    color: #cccccc;
    spacing: 7px;
    font-size: 8.5pt;
}

QCheckBox:hover, QRadioButton:hover {
    color: #ffffff;
}

QCheckBox::indicator, QRadioButton::indicator {
    width: 15px;
    height: 15px;
    border: 1px solid #3c3c3c;
    border-radius: 3px;
    background: #313131;
}

QRadioButton::indicator {
    border-radius: 7px;
}

QCheckBox::indicator:hover, QRadioButton::indicator:hover {
    border-color: #0078d4;
    background: #383838;
}

QCheckBox::indicator:checked {
    background-color: #0078d4;
    border-color: #0078d4;
}

QRadioButton::indicator:checked {
    background-color: #0078d4;
    border: 3px solid #313131;
}

/* Progress Bars */
QProgressBar {
    background-color: #252526;
    border: 1px solid #333333;
    border-radius: 4px;
    text-align: center;
    color: #ffffff;
    font-size: 8pt;
    font-weight: 600;
    min-height: 12px;
    max-height: 12px;
}

QProgressBar::chunk {
    background-color: #0078d4;
    border-radius: 3px;
}

/* Terminal & Log Text Areas */
QTextEdit, QPlainTextEdit {
    background-color: #181818;
    color: #cccccc;
    border: 1px solid #2b2b2b;
    border-radius: 4px;
    padding: 7px;
    font-family: 'Cascadia Code', 'Consolas', 'Google Sans Mono', monospace;
    font-size: 8pt;
    selection-background-color: #264f78;
    selection-color: #ffffff;
}

/* Lists and Tables */
QListWidget, QTableWidget, QTableView, QTreeView {
    background-color: #181818;
    color: #cccccc;
    border: 1px solid #2b2b2b;
    border-radius: 4px;
    gridline-color: #252526;
    selection-background-color: #04395e;
    selection-color: #ffffff;
    outline: none;
}

QListWidget::item, QTableWidget::item {
    padding: 4px 6px;
    border-radius: 3px;
}

QListWidget::item:hover, QTableWidget::item:hover {
    background-color: #2a2d2e;
}

QListWidget::item:selected, QTableWidget::item:selected {
    background-color: #04395e;
    color: #ffffff;
}

QHeaderView::section {
    background-color: #252526;
    color: #9d9d9d;
    border: none;
    border-bottom: 1px solid #2b2b2b;
    border-right: 1px solid #2b2b2b;
    padding: 5px 6px;
    font-weight: 600;
    font-size: 8pt;
}

QTableCornerButton::section {
    background-color: #252526;
    border: none;
}

/* Antigravity / VS Code Scrollbars */
QScrollBar:vertical {
    background: transparent;
    width: 8px;
    margin: 0;
}

QScrollBar::handle:vertical {
    background: rgba(121, 121, 121, 0.35);
    min-height: 24px;
    border-radius: 3px;
}

QScrollBar::handle:vertical:hover {
    background: rgba(100, 100, 100, 0.65);
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
    background: rgba(121, 121, 121, 0.35);
    min-width: 24px;
    border-radius: 3px;
}

QScrollBar::handle:horizontal:hover {
    background: rgba(100, 100, 100, 0.65);
}

QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal,
QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
    width: 0px;
    background: none;
}

/* Tabs */
QTabWidget::pane {
    border: 1px solid #333333;
    background-color: #252526;
    border-radius: 4px;
}

QTabBar::tab {
    background: #181818;
    color: #969696;
    padding: 6px 14px;
    border-top-left-radius: 4px;
    border-top-right-radius: 4px;
    margin-right: 2px;
    border: 1px solid #2b2b2b;
    border-bottom: none;
    font-weight: 500;
}

QTabBar::tab:selected {
    background: #252526;
    color: #ffffff;
    font-weight: 600;
    border-bottom: 2px solid #0078d4;
}

QTabBar::tab:hover:!selected {
    background: #1f1f1f;
    color: #cccccc;
}

/* Tooltips */
QToolTip {
    background-color: #252526;
    color: #cccccc;
    border: 1px solid #3c3c3c;
    border-radius: 4px;
    padding: 4px 8px;
    font-size: 8pt;
}

/* Antigravity Secondary PushButton */
QPushButton {
    background: #313131;
    color: #cccccc;
    border: 1px solid #3c3c3c;
    border-radius: 4px;
    padding: 5px 12px;
    font-size: 8.5pt;
    font-weight: 500;
}

QPushButton:hover {
    background: #3c3c3c;
    border-color: #4a4a4a;
    color: #ffffff;
}

QPushButton:pressed {
    background: #252526;
}

QPushButton:disabled {
    background: #222222;
    color: #656565;
    border-color: #2d2d2d;
}

/* Primary role button */
QPushButton[role=primary] {
    background: #0078d4;
    color: #ffffff;
    border: 1px solid #0078d4;
    border-radius: 4px;
    padding: 6px 16px;
    font-weight: 600;
    font-size: 8.5pt;
}

QPushButton[role=primary]:hover {
    background: #006cc1;
    border-color: #006cc1;
}

QPushButton[role=primary]:pressed {
    background: #005a9e;
}

/* Secondary role button */
QPushButton[role=secondary] {
    background: #313131;
    color: #cccccc;
    border: 1px solid #3c3c3c;
    border-radius: 4px;
    padding: 5px 12px;
    font-weight: 500;
}

QPushButton[role=secondary]:hover {
    background: #3c3c3c;
    border-color: #4a4a4a;
    color: #ffffff;
}
"""

START_BTN = """
QPushButton {
    background-color: #0078d4;
    color: #ffffff;
    border: 1px solid #0078d4;
    border-radius: 4px;
    padding: 6px 18px;
    font-weight: 600;
    font-size: 8.5pt;
}
QPushButton:hover {
    background-color: #006cc1;
    border-color: #006cc1;
}
QPushButton:pressed {
    background-color: #005a9e;
    border-color: #005a9e;
}
QPushButton:disabled {
    background-color: #222222;
    color: #656565;
    border: 1px solid #2d2d2d;
}
"""

STOP_BTN = """
QPushButton {
    background-color: #c42b1c;
    color: #ffffff;
    border: 1px solid #c42b1c;
    border-radius: 4px;
    padding: 6px 16px;
    font-weight: 600;
    font-size: 8.5pt;
}
QPushButton:hover {
    background-color: #b3261e;
    border-color: #b3261e;
}
QPushButton:pressed {
    background-color: #9c1f17;
    border-color: #9c1f17;
}
QPushButton:disabled {
    background-color: #222222;
    color: #656565;
    border: 1px solid #2d2d2d;
}
"""

NAV_STYLE = """
QPushButton {
    background: transparent;
    color: #858585;
    border: none;
    border-left: 2px solid transparent;
    border-radius: 0px;
    padding: 6px 10px;
    text-align: left;
    font-size: 8.5pt;
    font-weight: 500;
    margin: 1px 0px;
}
QPushButton:hover {
    background: #2a2d2e;
    color: #ffffff;
}
QPushButton[active="true"] {
    background: #252526;
    color: #ffffff;
    border-left: 2px solid #ffffff;
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
        if color in ("#7eb8f7", "#60a5fa", "#38bdf8"):
            c = "#ffffff"
        elif color in ("#556", "#64748b", "#94a3b8", "#555", "#666"):
            c = "#858585"
        else:
            c = color
        w.setStyleSheet(f"color: {c};")
    return w


def hline():
    f = QFrame()
    f.setFrameShape(QFrame.HLine)
    f.setStyleSheet("background-color: #2b2b2b; border: none; height: 1px; max-height: 1px;")
    return f


def notify(parent, title, message, critical=False, file_path=None, folder_path=None):
    """Show a result without blocking the Qt event loop with exec()."""
    box = QMessageBox(QMessageBox.Critical if critical else QMessageBox.Information,
                      title, message, parent=parent)
    box.setAttribute(Qt.WA_DeleteOnClose)
    box.setStyleSheet("""
        QMessageBox {
            background-color: #252526;
            color: #cccccc;
        }
        QLabel {
            color: #cccccc;
            font-size: 8.5pt;
        }
        QPushButton {
            background: #313131;
            color: #cccccc;
            border: 1px solid #3c3c3c;
            border-radius: 4px;
            padding: 5px 14px;
            font-weight: 500;
        }
        QPushButton:hover {
            background: #0078d4;
            color: #ffffff;
            border-color: #0078d4;
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

import os
from PySide6.QtGui import QPalette, QColor
from PySide6.QtWidgets import (
    QLabel, QFrame, QMessageBox, QDialog, QVBoxLayout, QHBoxLayout,
    QPushButton, QTextEdit, QStyle
)
from PySide6.QtCore import Qt

def configure_app(app):
    app.setStyle("Fusion")
    font = app.font()
    font.setPointSizeF(font.pointSizeF())
    app.setFont(font)
    palette = QPalette()
    colors = {
        "Window": "#20242c", "WindowText": "#edf0f5",
        "Base": "#161a21", "AlternateBase": "#292f3a",
        "ToolTipBase": "#edf0f5", "ToolTipText": "#161a21",
        "Text": "#edf0f5", "Button": "#303846",
        "ButtonText": "#edf0f5", "BrightText": "#ffffff",
        "Highlight": "#347bd1", "HighlightedText": "#ffffff",
        "PlaceholderText": "#9ba8ba", "Link": "#79b7ff",
    }
    for role, color in colors.items():
        palette.setColor(getattr(QPalette.ColorRole, role), QColor(color))
    for role in (QPalette.ColorRole.Text, QPalette.ColorRole.WindowText, QPalette.ColorRole.ButtonText):
        palette.setColor(QPalette.ColorGroup.Disabled, role, QColor("#818b9b"))
    app.setPalette(palette)

DARK_STYLE = """
        QPushButton[role=primary] { background: #3c9b5f; color: white; border: 1px solid #55b876; border-radius: 5px; padding: 6px 14px; font-weight: 600; }
        QPushButton[role=primary]:hover { background: #49ad6d; }
        QPushButton[role=secondary] { background: #b8753c; color: white; border: 1px solid #d39152; border-radius: 5px; padding: 5px 11px; }
        QPushButton[role=secondary]:hover { background: #ca874a; }
        QPushButton:disabled { color: #818b9b; background: #363d49; border-color: #4a5361; }
    """

START_BTN = "QPushButton { background:#3c9b5f; color:white; border:1px solid #55b876; border-radius:5px; padding:6px 14px; font-weight:600; } QPushButton:hover { background:#49ad6d; } QPushButton:disabled { color:#818b9b; background:#363d49; border-color:#4a5361; }"
STOP_BTN = "QPushButton { background:#b8753c; color:white; border:1px solid #d39152; border-radius:5px; padding:5px 11px; } QPushButton:hover { background:#ca874a; } QPushButton:disabled { color:#818b9b; background:#363d49; border-color:#4a5361; }"


def lbl(text, bold=False, color=None):
    w = QLabel(text)
    w.setWordWrap(True)
    if bold:
        f = w.font(); f.setBold(True); w.setFont(f)
    if color:
        w.setStyleSheet(f"color:{color};")
    return w


def hline():
    f = QFrame()
    f.setFrameShape(QFrame.HLine)
    f.setStyleSheet("color:#2e3340;")
    return f


class NotificationDialog(QDialog):
    """
    Sleek, responsive notification dialog with auto-scroll for long messages
    and proper keyboard controls (Esc/Enter) so the user never gets stuck.
    """
    def __init__(self, parent=None, title="Notification", message="", critical=False,
                 file_path=None, folder_path=None, details=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self.setWindowFlags(Qt.Dialog | Qt.WindowTitleHint | Qt.WindowCloseButtonHint)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setMinimumWidth(440)
        self.setMaximumWidth(620)
        self.setMaximumHeight(520)

        self.setStyleSheet("""
            QDialog {
                background: #1e222a;
                color: #edf0f5;
                font-family: 'Segoe UI', Tahoma, sans-serif;
            }
            QPushButton {
                background: #334155;
                color: #edf0f5;
                font-size: 9pt;
                font-weight: 500;
                padding: 6px 16px;
                border-radius: 5px;
                border: 1px solid #475569;
            }
            QPushButton:hover {
                background: #475569;
                color: #ffffff;
            }
            QPushButton:pressed {
                background: #1e293b;
            }
            QPushButton#ok_btn {
                background: #347bd1;
                color: #ffffff;
                border: 1px solid #4a90e2;
                font-weight: 600;
                min-width: 75px;
            }
            QPushButton#ok_btn:hover {
                background: #4188de;
            }
            QPushButton#ok_btn:pressed {
                background: #2b6cb8;
            }
        """)

        # Separate summary from details if not explicitly passed
        if details is None:
            if "\n\n" in message:
                summary, details = message.split("\n\n", 1)
            else:
                lines = message.splitlines()
                if len(lines) > 5:
                    summary = lines[0]
                    details = "\n".join(lines[1:])
                else:
                    summary = message
                    details = None
        else:
            summary = message

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(12)

        # Header: Icon + Summary text
        header_layout = QHBoxLayout()
        header_layout.setSpacing(14)
        header_layout.setAlignment(Qt.AlignTop)

        icon_type = QStyle.SP_MessageBoxCritical if critical else QStyle.SP_MessageBoxInformation
        icon_pixmap = self.style().standardIcon(icon_type).pixmap(36, 36)
        icon_lbl = QLabel()
        icon_lbl.setPixmap(icon_pixmap)
        icon_lbl.setFixedSize(36, 36)
        header_layout.addWidget(icon_lbl)

        summary_lbl = QLabel(summary.strip())
        summary_lbl.setWordWrap(True)
        summary_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
        summary_lbl.setStyleSheet("color: #edf0f5; font-size: 10pt; font-weight: 500; line-height: 1.4;")
        header_layout.addWidget(summary_lbl, 1)
        layout.addLayout(header_layout)

        # Details scroll box if any
        if details and details.strip():
            details_edit = QTextEdit()
            details_edit.setReadOnly(True)
            details_edit.setPlainText(details.strip())
            details_edit.setMinimumHeight(110)
            details_edit.setMaximumHeight(260)
            details_edit.setStyleSheet("""
                QTextEdit {
                    background-color: #161a21;
                    color: #edf0f5;
                    border: 1px solid #374151;
                    border-radius: 6px;
                    padding: 8px;
                    font-family: Consolas, 'Segoe UI', monospace;
                    font-size: 9pt;
                }
            """)
            layout.addWidget(details_edit)

        # Action and Dismiss buttons
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)

        if file_path:
            file_str = str(file_path).lower()
            btn_label = "Open Excel" if file_str.endswith((".xlsx", ".xls", ".xlsm", ".csv")) else "Open the file"
            btn_file = QPushButton(btn_label)
            btn_file.setCursor(Qt.PointingHandCursor)
            btn_file.clicked.connect(lambda _=False, p=file_path: os.startfile(p) if os.path.exists(p) else None)
            btn_layout.addWidget(btn_file)

        if folder_path:
            btn_folder = QPushButton("Open Folder")
            btn_folder.setCursor(Qt.PointingHandCursor)
            btn_folder.clicked.connect(lambda _=False, p=folder_path: os.startfile(p) if os.path.exists(p) else None)
            btn_layout.addWidget(btn_folder)

        btn_layout.addStretch(1)

        ok_btn = QPushButton("OK")
        ok_btn.setObjectName("ok_btn")
        ok_btn.setCursor(Qt.PointingHandCursor)
        ok_btn.setDefault(True)
        ok_btn.setAutoDefault(True)
        ok_btn.clicked.connect(self.accept)
        btn_layout.addWidget(ok_btn)

        layout.addLayout(btn_layout)

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key_Escape, Qt.Key_Return, Qt.Key_Enter):
            self.accept()
            return
        super().keyPressEvent(event)


def notify(parent, title, message, critical=False, file_path=None, folder_path=None, details=None):
    """Show a result without blocking the Qt event loop with exec()."""
    dlg = NotificationDialog(
        parent=parent,
        title=title,
        message=message,
        critical=critical,
        file_path=file_path,
        folder_path=folder_path,
        details=details
    )
    active = getattr(parent, "_active_notifications", None)
    if active is None:
        active = parent._active_notifications = []
    active.append(dlg)
    dlg.finished.connect(lambda: active.remove(dlg) if dlg in active else None)
    dlg.open()
    return dlg


# ──────────────────────────────────────────────────────────────
#  TOOL 1 — PITC Downloader (CDS10)
#  All report types · OTP login · Month+Amount filter · Highlight · Sr. No. Stamp
# ──────────────────────────────────────────────────────────────

NAV_STYLE = """
QPushButton {
    background: transparent; color: #b8c0cc;
    border: none; border-left: 3px solid transparent;
    border-radius: 0; padding: 10px 12px; text-align: left;
}
QPushButton:hover { background: #292f3a; color: #edf0f5; }
QPushButton[active="true"] {
    background: #303846; color: #79b7ff;
    border-left-color: #347bd1; font-weight: 600;
}
"""

import os
from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QLabel, QPushButton, QVBoxLayout, QHBoxLayout,
    QStackedWidget, QFrame, QLineEdit, QSpinBox, QSizePolicy, QScrollArea,
    QAbstractScrollArea, QBoxLayout, QDialog, QApplication
)
from PySide6.QtCore import Qt, QSize, QTimer, QSettings
from PySide6.QtGui import QPixmap

from pitcsuite.icons import app_icon, pdf_merger_icon, get_icon
from pitcsuite.ui_helpers import NAV_STYLE
from pitcsuite.tools.pitc_downloader import PITCPanel
from pitcsuite.tools.extract_page import ExtractPagePanel
from pitcsuite.tools.bill_downloader import BillSoftPanel
from pitcsuite.tools.payment_extract_88l import PaymentExtract88LPanel
from pitcsuite.tools.pdf_merger import PDFMerger
from pitcsuite.tools.file_checker import SrNoCheckerPanel
from pitcsuite.tools.file_renamer import FileRenamerPanel
from pitcsuite.tools.double_cr import DoubleCRPanel
from pitcsuite.tools.settings import SettingsPanel
from pitcsuite.auto_updater import (
    check_for_updates, get_current_version, get_release_notes,
    UpToDateDialog, UpdateAvailableDialog, start_download, show_post_update_dialog
)
from pitcsuite.tools.integrated_tools import (
    CP22TCheckerPanel, BillScraperPanel, FileMoverPanel,
    CP52PostingCheckerPanel, CP136PostingCheckerPanel, CP139PostingCheckerPanel, CP21cConverterPanel
)


TOOLS = [
    ("PITC Report Downloader",        PITCPanel, "pitc_down"),
    ("Extract Pages from PITC Reports", ExtractPagePanel, "extract_pages"),
    ("E/Bills Downloader",            BillSoftPanel, "ebills"),
    ("Bill Scrapper",                  BillScraperPanel, "bill_scrap"),
    ("Payment Extractor from 88L",    PaymentExtract88LPanel, "payment_ext"),
    ("CP 52 Posting to PDF",          CP52PostingCheckerPanel, "cp52"),
    ("CP136 & 137 Posting Checker",   CP136PostingCheckerPanel, "cp136"),
    ("CP 139 Posting Checker",        CP139PostingCheckerPanel, "cp139"),
    ("CP 22T Posting Checker",        CP22TCheckerPanel, "cp22t"),
    ("CP21c Converter",               CP21cConverterPanel, "cp21c"),
    ("Double CR Checker",             DoubleCRPanel, "double_cr"),
    ("PDF Merger",                    PDFMerger, "pdf_merger"),
    ("File Mover",                     FileMoverPanel, "file_mover"),
    ("File Renamer",                    FileRenamerPanel, "file_renamer"),
    ("Settings",                     SettingsPanel, "settings"),
]


class AboutDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("About PITC Suite")
        self.setFixedSize(400, 260)
        self.setModal(True)
        self.setWindowFlags(Qt.Dialog | Qt.WindowTitleHint | Qt.WindowCloseButtonHint)

        self.setStyleSheet("""
            QDialog {
                background: #20242c;
                color: #edf0f5;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(12)

        title_lbl = QLabel("PITC Suite")
        title_lbl.setStyleSheet("color: #7eb8f7; font-size: 14pt; font-weight: bold;")
        title_lbl.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_lbl)

        ver_lbl = QLabel(f"Version {get_current_version()}")
        ver_lbl.setStyleSheet("color: #9ba8ba; font-size: 9pt;")
        ver_lbl.setAlignment(Qt.AlignCenter)
        layout.addWidget(ver_lbl)

        dev_lbl = QLabel("Developed by TN Developers Co.\nEmail: tahir.naqash4@gmail.com\nContact: +923339962474")
        dev_lbl.setStyleSheet("color: #edf0f5; font-size: 10pt;")
        dev_lbl.setAlignment(Qt.AlignCenter)
        layout.addWidget(dev_lbl)

        layout.addStretch()

        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)
        btn_row.addStretch()

        whats_new_btn = QPushButton("✨ What's New")
        whats_new_btn.setCursor(Qt.PointingHandCursor)
        whats_new_btn.setStyleSheet("""
            QPushButton {
                background: #1e3a5f;
                color: #60a5fa;
                font-size: 8.5pt;
                font-weight: 600;
                padding: 6px 14px;
                border-radius: 5px;
                border: 1px solid #2563eb;
            }
            QPushButton:hover {
                background: #2563eb;
                color: #ffffff;
            }
        """)
        whats_new_btn.clicked.connect(self._show_whats_new)
        btn_row.addWidget(whats_new_btn)

        close_btn = QPushButton("Close")
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.setStyleSheet("""
            QPushButton {
                background: #334155;
                color: #e2e8f0;
                font-size: 9pt;
                padding: 6px 16px;
                border-radius: 5px;
                border: none;
            }
            QPushButton:hover {
                background: #475569;
            }
        """)
        close_btn.clicked.connect(self.accept)
        btn_row.addWidget(close_btn)
        btn_row.addStretch()
        layout.addLayout(btn_row)

    def _show_whats_new(self):
        notes = get_release_notes(get_current_version())
        dlg = UpToDateDialog(get_current_version(), notes, self)
        dlg.exec()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PITC Suite")
        self.setWindowIcon(app_icon())
        self.resize(790, 690)
        self.setMinimumSize(560, 520)
        root = QWidget()
        self.setCentralWidget(root)
        outer = QHBoxLayout(root)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)
        # Tool labels are visible when the application opens; the menu button
        # still toggles them into the compact icon-only view.
        self.nav_collapsed = False
        self.nav = QWidget()
        nav_layout = QVBoxLayout(self.nav)
        nav_layout.setContentsMargins(0, 6, 0, 6)
        nav_layout.setSpacing(0)
        self.nav_menu_btn = QPushButton(u"\u2630")
        self.nav_menu_btn.setToolTip("Show/hide tool labels")
        self.nav_menu_btn.clicked.connect(self.toggle_nav)
        nav_layout.addWidget(self.nav_menu_btn)
        self.nav_btns = []
        self.nav_names = []
        for i, (name, panel_class, icon_name) in enumerate(TOOLS):
            button = QPushButton(name)
            button.setIcon(get_icon(icon_name))
            button.setIconSize(QSize(20, 20))
            button.setToolTip(name)
            button.clicked.connect(lambda checked=False, idx=i: self.switch(idx))
            nav_layout.addWidget(button)
            self.nav_btns.append(button)
            self.nav_names.append(name)
        nav_layout.addStretch()



        self.update_btn = QPushButton("Check for Updates")
        self.update_btn.setStyleSheet("""
            QPushButton {
                background: #334155;
                color: #e2e8f0;
                font-size: 7.5pt;
                padding: 4px;
                border-radius: 4px;
                border: none;
                margin: 4px 8px 8px 8px;
            }
            QPushButton:hover {
                background: #475569;
            }
        """)
        self.update_btn.clicked.connect(self._on_update_clicked)
        nav_layout.addWidget(self.update_btn)

        self._apply_nav_state()
        outer.addWidget(self.nav)
        content = QWidget()
        outer.addWidget(content, 1)
        layout = QVBoxLayout(content)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(6)
        heading = QLabel("PITC Suite")
        heading.setStyleSheet("color: #7eb8f7; font-size: 22pt; font-weight: bold; letter-spacing: 1px;")
        layout.addWidget(heading)
        self.stack = QStackedWidget()
        self.panels = []
        for name, panel_class, icon_name in TOOLS:
            panel = panel_class()
            self.panels.append(panel)
            if panel.layout():
                panel.layout().setContentsMargins(0, 6, 0, 0)
                panel.layout().setAlignment(Qt.AlignTop)
            for button in panel.findChildren(QPushButton):
                if button.property("role") is None and not button.styleSheet():
                    button.setProperty("role", "secondary")
                button.setMinimumWidth(0)
                button.setMaximumWidth(16777215)
            for edit in panel.findChildren(QLineEdit):
                if isinstance(edit.parentWidget(), QSpinBox):
                    continue
                edit.setMinimumWidth(70)
                edit.setMaximumWidth(16777215)
                edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            scroll.setFrameShape(QFrame.NoFrame)
            scroll.setAlignment(Qt.AlignLeft | Qt.AlignTop)
            scroll.setWidget(panel)
            # Hide the controls without disabling wheel, touchpad or keyboard scrolling.
            for area in [scroll, *panel.findChildren(QAbstractScrollArea)]:
                area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
                area.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            self.stack.addWidget(scroll)
        layout.addWidget(self.stack, 1)
        footer_layout = QHBoxLayout()
        footer_layout.setContentsMargins(0, 4, 0, 0)
        footer_layout.setSpacing(10)
        footer_layout.addStretch()

        about_btn = QPushButton("About")
        about_btn.setCursor(Qt.PointingHandCursor)
        about_btn.setStyleSheet("""
            QPushButton {
                background: #334155;
                color: #e2e8f0;
                font-size: 7.5pt;
                padding: 4px 12px;
                border-radius: 4px;
                border: none;
            }
            QPushButton:hover {
                background: #475569;
            }
        """)
        about_btn.clicked.connect(self._show_about_dialog)
        footer_layout.addWidget(about_btn)

        footer_layout.addStretch()
        layout.addLayout(footer_layout)
        self.switch(0)

        # Check for updates in the background after 1 second
        QTimer.singleShot(1000, lambda: check_for_updates(self, silent=True))

        # Check if app was just updated to show What's New notes
        self._check_post_update()

    def _show_about_dialog(self):
        dlg = AboutDialog(self)
        dlg.exec()

    def switch(self, idx):
        self.stack.setCurrentIndex(idx)
        for i, button in enumerate(self.nav_btns):
            button.setProperty("active", "true" if i == idx else "false")
            button.style().unpolish(button)
            button.style().polish(button)

    def toggle_nav(self):
        self.nav_collapsed = not self.nav_collapsed
        self._apply_nav_state()

    def notify_update_available(self, latest_tag=None, notes=None):
        self._has_update = True
        self._latest_tag = latest_tag
        self._latest_notes = notes
        self.update_btn.setStyleSheet("""
            QPushButton {
                background: #22c55e;
                color: #ffffff;
                font-size: 7.5pt;
                font-weight: bold;
                padding: 4px;
                border-radius: 4px;
                border: none;
                margin: 4px 8px 8px 8px;
            }
            QPushButton:hover {
                background: #16a34a;
            }
        """)
        self._apply_nav_state()

    def _on_update_clicked(self):
        if getattr(self, "_has_update", False) and getattr(self, "_latest_tag", None):
            notes = getattr(self, "_latest_notes", None) or []
            dlg = UpdateAvailableDialog(get_current_version(), self._latest_tag, notes, self)
            if dlg.exec() == QDialog.Accepted:
                start_download(self, self._latest_tag, notes)
        else:
            check_for_updates(self, silent=False)

    def _check_post_update(self):
        settings = QSettings("TNDevelopers", "PITCSuite")
        last_ver = settings.value("last_run_version", "")
        current_ver = get_current_version()
        if last_ver and last_ver != current_ver:
            settings.setValue("last_run_version", current_ver)
            QTimer.singleShot(1200, lambda: show_post_update_dialog(self, current_ver))
        elif not last_ver:
            settings.setValue("last_run_version", current_ver)

    def _apply_nav_state(self):
        self.nav.setFixedWidth(44 if self.nav_collapsed else 215)
        
        has_update = getattr(self, "_has_update", False)
        if self.nav_collapsed:
            self.update_btn.setText("!" if has_update else u"\U0001F504")
        else:
            self.update_btn.setText("Update Available!" if has_update else "Check for Updates")
        self.update_btn.setToolTip("Check for Updates")
        
        for button, name in zip(self.nav_btns, self.nav_names):
            if not button.icon().isNull():
                button.setText("" if self.nav_collapsed else name)
            else:
                button.setText(name.split()[0] if self.nav_collapsed else name)
            button.setStyleSheet(NAV_STYLE + (
                "\nQPushButton { padding:10px 0; text-align:center; }"
                if self.nav_collapsed else ""))

    def closeEvent(self, event):
        # Gracefully attempt to stop any background workers and drivers
        for panel in self.panels:
            for method in ['cancel', 'halt', 'stop', '_stop']:
                if hasattr(panel, method):
                    try:
                        getattr(panel, method)()
                    except:
                        pass
        super().closeEvent(event)

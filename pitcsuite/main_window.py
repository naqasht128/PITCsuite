import os
from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QLabel, QPushButton, QVBoxLayout, QHBoxLayout,
    QStackedWidget, QFrame, QLineEdit, QSpinBox, QSizePolicy, QScrollArea,
    QAbstractScrollArea, QBoxLayout, QDialog, QApplication
)
from PySide6.QtCore import Qt, QSize, QTimer, QSettings, QEvent, QObject
from PySide6.QtGui import QPixmap

from pitcsuite.icons import app_icon, pdf_merger_icon, get_icon
from pitcsuite.ui_helpers import NAV_STYLE
from pitcsuite.tools.pitc_downloader import PITCPanel
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
from pitcsuite.license_manager import (
    load_local_license, get_device_id, format_plan_type,
    format_expiry_date, LicenseGateDialog
)
from pitcsuite.tools.integrated_tools import (
    CP22TCheckerPanel, BillScraperPanel, FileMoverPanel,
    CP52PostingCheckerPanel, CP136PostingCheckerPanel, CP139PostingCheckerPanel, CP21cConverterPanel,
    AssessmentRealizationPanel
)


TOOLS = [
    ("PITC Report Downloader",        PITCPanel, "pitc_down"),
    ("E/Bills Downloader",            BillSoftPanel, "ebills"),
    ("Bill Scrapper",                  BillScraperPanel, "bill_scrap"),
    ("Payment Extractor from 88L",    PaymentExtract88LPanel, "payment_ext"),
    ("CP 52 Posting to PDF",          CP52PostingCheckerPanel, "cp52"),
    ("CP 136 & 137 Posting Checker",   CP136PostingCheckerPanel, "cp136"),
    ("CP 139 Posting Check/PDF",      CP139PostingCheckerPanel, "cp139"),
    ("CP 22T Posting Checker",        CP22TCheckerPanel, "cp22t"),
    ("CP21c Converter",               CP21cConverterPanel, "cp21c"),
    ("Double CR Checker",             DoubleCRPanel, "double_cr"),
    ("Govt Assessment Realization",   AssessmentRealizationPanel, "assessment_realization"),
    ("PDF Merger",                    PDFMerger, "pdf_merger"),
    ("File Mover",                     FileMoverPanel, "file_mover"),
    ("File Renamer",                    FileRenamerPanel, "file_renamer"),
    ("Settings",                     SettingsPanel, "settings"),
]


class AboutDialog(QDialog):
    """Sleek modal dialog showing application, active license, and developer information."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("About PITC Suite")
        self.setFixedSize(480, 480)
        self.setModal(True)
        self.setWindowFlags(Qt.Dialog | Qt.WindowTitleHint | Qt.WindowCloseButtonHint)

        self.setStyleSheet("""
            QDialog {
                background: #1e1e1e;
                color: #edf0f5;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 18, 22, 18)
        layout.setSpacing(10)

        # App Header
        title_lbl = QLabel("PITC Suite")
        title_lbl.setStyleSheet("color: #7eb8f7; font-size: 15pt; font-weight: bold;")
        layout.addWidget(title_lbl)

        ver_lbl = QLabel(f"Version {get_current_version()}")
        ver_lbl.setStyleSheet("color: #9ba8ba; font-size: 9pt;")
        layout.addWidget(ver_lbl)

        # 1. License Information Card
        lic = getattr(parent, "active_license", None) or load_local_license()
        serial_key = lic.get("serial_key", "Not Activated")
        plan_str = format_plan_type(lic.get("plan_type"))
        expiry_str = format_expiry_date(lic.get("expires_at"))

        lic_card = QFrame()
        lic_card.setStyleSheet("""
            QFrame {
                background: #252526;
                border: 1px solid #3c3c3c;
                border-radius: 6px;
            }
            QLabel {
                border: none;
                background: transparent;
            }
        """)
        lic_layout = QVBoxLayout(lic_card)
        lic_layout.setContentsMargins(14, 12, 14, 12)
        lic_layout.setSpacing(6)

        hdr_lbl = QLabel("LICENSE INFORMATION")
        hdr_lbl.setStyleSheet("color: #858585; font-size: 8pt; font-weight: bold;")
        lic_layout.addWidget(hdr_lbl)

        key_row = QHBoxLayout()
        key_title = QLabel("Key: ")
        key_title.setStyleSheet("color: #ffffff; font-size: 9pt; font-weight: bold;")
        key_row.addWidget(key_title)

        key_val = QLabel(serial_key)
        key_val.setStyleSheet("color: #4ec9b0; font-family: Consolas, monospace; font-size: 9.5pt; font-weight: bold;")
        key_row.addWidget(key_val)
        key_row.addStretch()

        self.change_key_btn = QPushButton("🔑 Change Key")
        self.change_key_btn.setCursor(Qt.PointingHandCursor)
        self.change_key_btn.setStyleSheet("""
            QPushButton {
                background: #2d2d2d;
                color: #ffffff;
                font-size: 8pt;
                font-weight: bold;
                padding: 4px 8px;
                border: 1px solid #3c3c3c;
                border-radius: 4px;
            }
            QPushButton:hover {
                background: #37373d;
            }
        """)
        self.change_key_btn.clicked.connect(self._change_serial_key)
        key_row.addWidget(self.change_key_btn)

        self.copy_key_btn = QPushButton("📋 Copy")
        self.copy_key_btn.setCursor(Qt.PointingHandCursor)
        self.copy_key_btn.setStyleSheet("""
            QPushButton {
                background: #2d2d2d;
                color: #cccccc;
                font-size: 8pt;
                padding: 4px 8px;
                border: 1px solid #3c3c3c;
                border-radius: 4px;
            }
            QPushButton:hover {
                background: #37373d;
                color: #ffffff;
            }
        """)
        self.copy_key_btn.clicked.connect(self._copy_serial_key)
        key_row.addWidget(self.copy_key_btn)
        lic_layout.addLayout(key_row)

        hw_lbl = QLabel(f"Hardware ID: {get_device_id()}")
        hw_lbl.setStyleSheet("color: #858585; font-family: Consolas, monospace; font-size: 8pt;")
        lic_layout.addWidget(hw_lbl)

        plan_lbl = QLabel(f"Plan: {plan_str}")
        plan_lbl.setStyleSheet("color: #cccccc; font-size: 9pt;")
        lic_layout.addWidget(plan_lbl)

        status_color = "#89d185" if "remaining" in expiry_str or "today" in expiry_str else "#cccccc"
        status_lbl = QLabel(f"Status: {expiry_str}")
        status_lbl.setStyleSheet(f"color: {status_color}; font-size: 9pt; font-weight: 500;")
        lic_layout.addWidget(status_lbl)

        layout.addWidget(lic_card)

        # 2. Developer Information Card
        dev_card = QFrame()
        dev_card.setStyleSheet("""
            QFrame {
                background: #252526;
                border: 1px solid #3c3c3c;
                border-radius: 6px;
            }
            QLabel {
                border: none;
                background: transparent;
            }
        """)
        dev_layout = QVBoxLayout(dev_card)
        dev_layout.setContentsMargins(14, 12, 14, 12)
        dev_layout.setSpacing(4)

        dev_title = QLabel("Developed by TN Developers Co.")
        dev_title.setStyleSheet("color: #ffffff; font-size: 9pt; font-weight: bold;")
        dev_layout.addWidget(dev_title)

        email_lbl = QLabel("Email: tahir.naqash4@gmail.com")
        email_lbl.setStyleSheet("color: #cccccc; font-size: 9pt;")
        dev_layout.addWidget(email_lbl)

        contact_lbl = QLabel("Contact: 03339962474")
        contact_lbl.setStyleSheet("color: #cccccc; font-size: 9pt;")
        dev_layout.addWidget(contact_lbl)

        layout.addWidget(dev_card)
        layout.addStretch()

        # Bottom Button Row
        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)

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
        btn_row.addStretch()

        close_btn = QPushButton("Close")
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.setStyleSheet("""
            QPushButton {
                background: #334155;
                color: #e2e8f0;
                font-size: 9pt;
                padding: 6px 18px;
                border-radius: 5px;
                border: none;
            }
            QPushButton:hover {
                background: #475569;
            }
        """)
        close_btn.clicked.connect(self.accept)
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

    def _copy_serial_key(self):
        lic = getattr(self.parent(), "active_license", None) or load_local_license()
        key = lic.get("serial_key", "")
        if key:
            QApplication.clipboard().setText(key)
            self.copy_key_btn.setText("✅ Copied!")
            self.copy_key_btn.setStyleSheet("""
                QPushButton {
                    background: #2d2d2d;
                    color: #4ec9b0;
                    font-size: 8pt;
                    padding: 4px 8px;
                    border: 1px solid #4ec9b0;
                    border-radius: 4px;
                }
            """)
            QTimer.singleShot(2000, lambda: self._reset_copy_btn() if not self.isHidden() else None)

    def _reset_copy_btn(self):
        self.copy_key_btn.setText("📋 Copy")
        self.copy_key_btn.setStyleSheet("""
            QPushButton {
                background: #2d2d2d;
                color: #cccccc;
                font-size: 8pt;
                padding: 4px 8px;
                border: 1px solid #3c3c3c;
                border-radius: 4px;
            }
            QPushButton:hover {
                background: #37373d;
                color: #ffffff;
            }
        """)

    def _change_serial_key(self):
        parent = self.parent()
        self.accept()
        gate = LicenseGateDialog(parent)
        gate.exec()

    def _show_whats_new(self):
        notes = get_release_notes(get_current_version())
        dlg = UpToDateDialog(get_current_version(), notes, self)
        dlg.exec()


class FramelessScrollArea(QScrollArea):
    """
    Scroll area with invisible scroll bars that supports:
    - Vertical wheel scrolling (standard mouse wheel)
    - Horizontal wheel scrolling (Shift + Wheel, Alt + Wheel)
    - Touchpad 2-finger horizontal swipe and mouse tilt wheel
    - Middle-click mouse drag to pan smoothly in both directions
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWidgetResizable(True)
        self.setFrameShape(QFrame.NoFrame)
        self.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._is_panning = False
        self._pan_start = None
        self._h_start = 0
        self._v_start = 0

    def wheelEvent(self, event):
        # 1. Native horizontal delta (touchpad 2-finger horizontal swipe or tilt wheel)
        dx = event.angleDelta().x()
        pixel_dx = event.pixelDelta().x() if hasattr(event, "pixelDelta") else 0
        if pixel_dx != 0:
            self.horizontalScrollBar().setValue(self.horizontalScrollBar().value() - pixel_dx)
            event.accept()
            return
        elif dx != 0:
            self.horizontalScrollBar().setValue(self.horizontalScrollBar().value() - dx)
            event.accept()
            return

        # 2. Shift or Alt + Mouse Wheel -> horizontal scroll
        modifiers = event.modifiers()
        if modifiers & (Qt.ShiftModifier | Qt.AltModifier):
            dy = event.angleDelta().y()
            pixel_dy = event.pixelDelta().y() if hasattr(event, "pixelDelta") else 0
            if pixel_dy != 0:
                self.horizontalScrollBar().setValue(self.horizontalScrollBar().value() - pixel_dy)
                event.accept()
                return
            elif dy != 0:
                self.horizontalScrollBar().setValue(self.horizontalScrollBar().value() - dy)
                event.accept()
                return

        super().wheelEvent(event)

    def mousePressEvent(self, event):
        if event.button() == Qt.MiddleButton:
            self._is_panning = True
            pos = event.position().toPoint() if hasattr(event, "position") else event.pos()
            self._pan_start = pos
            self._h_start = self.horizontalScrollBar().value()
            self._v_start = self.verticalScrollBar().value()
            self.viewport().setCursor(Qt.ClosedHandCursor)
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._is_panning and self._pan_start is not None:
            pos = event.position().toPoint() if hasattr(event, "position") else event.pos()
            delta = pos - self._pan_start
            self.horizontalScrollBar().setValue(self._h_start - delta.x())
            self.verticalScrollBar().setValue(self._v_start - delta.y())
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self._is_panning and event.button() == Qt.MiddleButton:
            self._is_panning = False
            self._pan_start = None
            self.viewport().unsetCursor()
            event.accept()
            return
        super().mouseReleaseEvent(event)


class ScrollEventFilter(QObject):
    """
    Event filter installed on child widgets and panels to ensure that:
    - Shift + Wheel (and Alt + Wheel)
    - Touchpad horizontal swipe
    - Middle-click drag
    are forwarded to the outer FramelessScrollArea instead of being consumed by child widgets.
    """
    def __init__(self, scroll_area: FramelessScrollArea):
        super().__init__(scroll_area)
        self.scroll_area = scroll_area

    def eventFilter(self, watched, event):
        etype = event.type()
        if etype == QEvent.Wheel:
            # If horizontal intent: Shift or Alt is held, or horizontal delta is present
            if (event.modifiers() & (Qt.ShiftModifier | Qt.AltModifier)) or event.angleDelta().x() != 0:
                self.scroll_area.wheelEvent(event)
                return True
        elif etype == QEvent.MouseButtonPress and event.button() == Qt.MiddleButton:
            self.scroll_area.mousePressEvent(event)
            return True
        elif etype == QEvent.MouseMove and getattr(self.scroll_area, '_is_panning', False):
            self.scroll_area.mouseMoveEvent(event)
            return True
        elif etype == QEvent.MouseButtonRelease and getattr(self.scroll_area, '_is_panning', False) and event.button() == Qt.MiddleButton:
            self.scroll_area.mouseReleaseEvent(event)
            return True
        elif etype == QEvent.ChildAdded and hasattr(event, "child") and event.child().isWidgetType():
            event.child().installEventFilter(self)
        return False


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
            scroll = FramelessScrollArea()
            scroll.setWidget(panel)
            filt = ScrollEventFilter(scroll)
            panel.installEventFilter(filt)
            for child in panel.findChildren(QWidget):
                child.installEventFilter(filt)
                if isinstance(child, QAbstractScrollArea):
                    child.viewport().installEventFilter(filt)
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

    def notify_update_available(self, latest_tag=None, notes=None, download_url=None, file_size=0):
        self._has_update = True
        self._latest_tag = latest_tag
        self._latest_notes = notes
        self._download_url = download_url
        self._latest_file_size = file_size
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
            dl_url = getattr(self, "_download_url", None)
            f_size = getattr(self, "_latest_file_size", 0)
            dlg = UpdateAvailableDialog(get_current_version(), self._latest_tag, notes, self)
            if dlg.exec() == QDialog.Accepted:
                start_download(self, self._latest_tag, notes, download_url=dl_url, file_size=f_size)
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

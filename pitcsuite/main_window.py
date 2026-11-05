import os
from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QLabel, QPushButton, QVBoxLayout, QHBoxLayout,
    QStackedWidget, QFrame, QLineEdit, QSpinBox, QSizePolicy, QScrollArea,
    QAbstractScrollArea, QDialog, QApplication, QGraphicsOpacityEffect
)
from PySide6.QtCore import (
    Qt, QSize, QTimer, QSettings, QEvent, QObject, QPropertyAnimation,
    QEasingCurve, QParallelAnimationGroup, QAbstractAnimation, QByteArray
)
from PySide6.QtGui import QPixmap, QIcon

from pitcsuite.icons import app_icon, pdf_merger_icon, get_icon
from pitcsuite.ui_helpers import NAV_STYLE
from pitcsuite.tools.pitc_downloader import PITCPanel
from pitcsuite.tools.bill_downloader import BillSoftPanel
from pitcsuite.tools.payment_extract_88l import PaymentExtract88LPanel
from pitcsuite.tools.pdf_merger import PDFMerger
from pitcsuite.tools.file_checker import SrNoCheckerPanel
from pitcsuite.tools.file_renamer import FileRenamerPanel
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
    ("PITC Report Downloader",        PITCPanel, "pitc_down", "Automated CDS10 reports, OTP login & filtering"),
    ("E/Bills Downloader",            BillSoftPanel, "ebills", "Batch download consumer electricity bills"),
    ("Bill Scrapper",                  BillScraperPanel, "bill_scrap", "Automated web scraper for bill data extraction"),
    ("Payment Extractor from 88L",    PaymentExtract88LPanel, "payment_ext", "Extract and parse payment data from 88L reports"),
    ("CP 52 Posting to PDF",          CP52PostingCheckerPanel, "cp52", "Process CP 52 posting records and generate PDF"),
    ("CP 136 & 137 Posting Checker",   CP136PostingCheckerPanel, "cp136", "Audit and verify CP 136 / 137 posting records"),
    ("CP 139 Posting Check/PDF",      CP139PostingCheckerPanel, "cp139", "Verify CP 139 records & generate PDF reports"),
    ("CP 22T Posting Checker",        CP22TCheckerPanel, "cp22t", "Posting verification and reconciliation"),
    ("CP21c Converter",               CP21cConverterPanel, "cp21c", "Convert and format CP21c records"),
    ("Govt Assessment Realization",   AssessmentRealizationPanel, "assessment_realization", "Extract government assessment & realization data"),
    ("PDF Merger",                    PDFMerger, "pdf_merger", "Merge, split, and stamp PDF documents"),
    ("File Mover",                     FileMoverPanel, "file_mover", "Organize and move files by batch pattern"),
    ("File Renamer",                    FileRenamerPanel, "file_renamer", "Batch rename files with custom rules"),
    ("Settings",                     SettingsPanel, "settings", "DISCO configuration & login credentials"),
]


class AboutDialog(QDialog):
    """Sleek modal dialog showing application, active license, and developer information."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("About PITC Suite")
        self.setFixedSize(500, 500)
        self.setModal(True)
        self.setWindowFlags(Qt.Dialog | Qt.WindowTitleHint | Qt.WindowCloseButtonHint)

        self.setStyleSheet("""
            QDialog {
                background: #0d1117;
                color: #f0f6fc;
                font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 22, 24, 20)
        layout.setSpacing(12)

        # App Header
        header_row = QHBoxLayout()
        header_row.setSpacing(10)
        
        title_vbox = QVBoxLayout()
        title_vbox.setSpacing(2)
        
        title_lbl = QLabel("PITC Suite")
        title_lbl.setStyleSheet("color: #60a5fa; font-size: 16pt; font-weight: bold; letter-spacing: 0.5px;")
        title_vbox.addWidget(title_lbl)

        ver_lbl = QLabel(f"Version {get_current_version()} · Windows Edition")
        ver_lbl.setStyleSheet("color: #94a3b8; font-size: 8.5pt;")
        title_vbox.addWidget(ver_lbl)
        
        header_row.addLayout(title_vbox)
        header_row.addStretch()
        layout.addLayout(header_row)

        # 1. License Information Card
        lic = getattr(parent, "active_license", None) or load_local_license()
        serial_key = lic.get("serial_key", "Not Activated")
        plan_str = format_plan_type(lic.get("plan_type"))
        expiry_str = format_expiry_date(lic.get("expires_at"))

        lic_card = QFrame()
        lic_card.setStyleSheet("""
            QFrame {
                background: #121722;
                border: 1px solid #1e2638;
                border-radius: 8px;
            }
            QLabel {
                border: none;
                background: transparent;
            }
        """)
        lic_layout = QVBoxLayout(lic_card)
        lic_layout.setContentsMargins(16, 14, 16, 14)
        lic_layout.setSpacing(8)

        hdr_lbl = QLabel("LICENSE & REGISTRATION")
        hdr_lbl.setStyleSheet("color: #64748b; font-size: 8pt; font-weight: bold; letter-spacing: 0.5px;")
        lic_layout.addWidget(hdr_lbl)

        key_row = QHBoxLayout()
        key_title = QLabel("Key:")
        key_title.setStyleSheet("color: #cbd5e1; font-size: 9pt; font-weight: 600;")
        key_row.addWidget(key_title)

        key_val = QLabel(serial_key)
        key_val.setStyleSheet("color: #38bdf8; font-family: 'Cascadia Code', Consolas, monospace; font-size: 9pt; font-weight: 600;")
        key_row.addWidget(key_val)
        key_row.addStretch()

        self.change_key_btn = QPushButton("🔑 Change")
        self.change_key_btn.setCursor(Qt.PointingHandCursor)
        self.change_key_btn.setStyleSheet("""
            QPushButton {
                background: #1e2638;
                color: #e2e8f0;
                font-size: 8pt;
                font-weight: 600;
                padding: 4px 10px;
                border: 1px solid #2e3a50;
                border-radius: 5px;
            }
            QPushButton:hover {
                background: #25334a;
                border-color: #3b82f6;
                color: #ffffff;
            }
        """)
        self.change_key_btn.clicked.connect(self._change_serial_key)
        key_row.addWidget(self.change_key_btn)

        self.copy_key_btn = QPushButton("📋 Copy")
        self.copy_key_btn.setCursor(Qt.PointingHandCursor)
        self.copy_key_btn.setStyleSheet("""
            QPushButton {
                background: #1e2638;
                color: #94a3b8;
                font-size: 8pt;
                padding: 4px 10px;
                border: 1px solid #2e3a50;
                border-radius: 5px;
            }
            QPushButton:hover {
                background: #25334a;
                color: #ffffff;
            }
        """)
        self.copy_key_btn.clicked.connect(self._copy_serial_key)
        key_row.addWidget(self.copy_key_btn)
        lic_layout.addLayout(key_row)

        hw_lbl = QLabel(f"Hardware ID:  {get_device_id()}")
        hw_lbl.setStyleSheet("color: #64748b; font-family: 'Cascadia Code', Consolas, monospace; font-size: 8pt;")
        lic_layout.addWidget(hw_lbl)

        info_row = QHBoxLayout()
        plan_lbl = QLabel(f"Plan: <b style='color:#f0f6fc;'>{plan_str}</b>")
        plan_lbl.setStyleSheet("color: #94a3b8; font-size: 8.5pt;")
        info_row.addWidget(plan_lbl)
        info_row.addStretch()

        status_color = "#34d399" if "remaining" in expiry_str or "today" in expiry_str or "Lifetime" in plan_str else "#94a3b8"
        status_lbl = QLabel(f"Status: <b style='color:{status_color};'>{expiry_str}</b>")
        status_lbl.setStyleSheet("font-size: 8.5pt;")
        info_row.addWidget(status_lbl)
        lic_layout.addLayout(info_row)

        layout.addWidget(lic_card)

        # 2. Developer Information Card
        dev_card = QFrame()
        dev_card.setStyleSheet("""
            QFrame {
                background: #121722;
                border: 1px solid #1e2638;
                border-radius: 8px;
            }
            QLabel {
                border: none;
                background: transparent;
            }
        """)
        dev_layout = QVBoxLayout(dev_card)
        dev_layout.setContentsMargins(16, 14, 16, 14)
        dev_layout.setSpacing(6)

        dev_title = QLabel("DEVELOPER SUPPORT")
        dev_title.setStyleSheet("color: #64748b; font-size: 8pt; font-weight: bold; letter-spacing: 0.5px;")
        dev_layout.addWidget(dev_title)

        org_lbl = QLabel("TN Developers Co.")
        org_lbl.setStyleSheet("color: #f0f6fc; font-size: 9.5pt; font-weight: 600;")
        dev_layout.addWidget(org_lbl)

        email_lbl = QLabel("Email:  <span style='color: #38bdf8;'>tahir.naqash4@gmail.com</span>")
        email_lbl.setStyleSheet("color: #94a3b8; font-size: 8.5pt;")
        dev_layout.addWidget(email_lbl)

        contact_lbl = QLabel("WhatsApp / Contact:  <span style='color: #34d399;'>03339962474</span>")
        contact_lbl.setStyleSheet("color: #94a3b8; font-size: 8.5pt;")
        dev_layout.addWidget(contact_lbl)

        layout.addWidget(dev_card)
        layout.addStretch()

        # Bottom Button Row
        btn_row = QHBoxLayout()
        btn_row.setSpacing(12)

        whats_new_btn = QPushButton("✨ What's New")
        whats_new_btn.setCursor(Qt.PointingHandCursor)
        whats_new_btn.setStyleSheet("""
            QPushButton {
                background: #172554;
                color: #60a5fa;
                font-size: 8.5pt;
                font-weight: 600;
                padding: 7px 16px;
                border-radius: 6px;
                border: 1px solid #1d4ed8;
            }
            QPushButton:hover {
                background: #1d4ed8;
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
                background: #1e2638;
                color: #e2e8f0;
                font-size: 8.5pt;
                padding: 7px 20px;
                border-radius: 6px;
                border: 1px solid #2e3a50;
            }
            QPushButton:hover {
                background: #28334a;
                color: #ffffff;
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
                    background: #142e23;
                    color: #34d399;
                    font-size: 8pt;
                    padding: 4px 10px;
                    border: 1px solid #059669;
                    border-radius: 5px;
                }
            """)
            QTimer.singleShot(2000, lambda: self._reset_copy_btn() if not self.isHidden() else None)

    def _reset_copy_btn(self):
        self.copy_key_btn.setText("📋 Copy")
        self.copy_key_btn.setStyleSheet("""
            QPushButton {
                background: #1e2638;
                color: #94a3b8;
                font-size: 8pt;
                padding: 4px 10px;
                border: 1px solid #2e3a50;
                border-radius: 5px;
            }
            QPushButton:hover {
                background: #25334a;
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
        self.setMinimumSize(560, 500)
        
        # 1. Window size & geometry restoration (700 x 720 initial, or last saved)
        self._init_window_size()

        root = QWidget()
        root.setObjectName("centralWidget")
        self.setCentralWidget(root)
        outer = QHBoxLayout(root)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # 2. Sleek Dark Sidebar Navigation
        self.nav_collapsed = False
        self.nav = QWidget()
        self.nav.setObjectName("navSidebar")
        self.nav.setStyleSheet("""
            QWidget#navSidebar {
                background-color: #111620;
                border-right: 1px solid #1e2638;
            }
        """)
        nav_layout = QVBoxLayout(self.nav)
        nav_layout.setContentsMargins(0, 8, 0, 8)
        nav_layout.setSpacing(1)

        # Sidebar Header: Toggle button & Branding
        nav_header = QHBoxLayout()
        nav_header.setContentsMargins(10, 4, 10, 8)
        nav_header.setSpacing(8)

        self.nav_menu_btn = QPushButton("☰")
        self.nav_menu_btn.setFixedSize(32, 32)
        self.nav_menu_btn.setToolTip("Toggle sidebar labels")
        self.nav_menu_btn.setCursor(Qt.PointingHandCursor)
        self.nav_menu_btn.setStyleSheet("""
            QPushButton {
                background: #18202e;
                color: #94a3b8;
                border: 1px solid #253043;
                border-radius: 6px;
                font-size: 11pt;
                padding: 0;
            }
            QPushButton:hover {
                background: #25334a;
                border-color: #3b82f6;
                color: #f8fafc;
            }
        """)
        self.nav_menu_btn.clicked.connect(self.toggle_nav)
        nav_header.addWidget(self.nav_menu_btn)

        self.nav_brand_lbl = QLabel("PITC Suite")
        self.nav_brand_lbl.setStyleSheet("color: #60a5fa; font-size: 11pt; font-weight: bold; letter-spacing: 0.5px;")
        nav_header.addWidget(self.nav_brand_lbl)
        nav_header.addStretch()
        nav_layout.addLayout(nav_header)

        # Separator line under sidebar header
        sep1 = QFrame()
        sep1.setFrameShape(QFrame.HLine)
        sep1.setStyleSheet("background-color: #1e2638; border: none; height: 1px; max-height: 1px; margin: 0 8px 6px 8px;")
        nav_layout.addWidget(sep1)

        # Tool buttons
        self.nav_btns = []
        self.nav_names = []
        for i, item in enumerate(TOOLS):
            name = item[0]
            icon_name = item[2]
            button = QPushButton(name)
            button.setIcon(get_icon(icon_name))
            button.setIconSize(QSize(18, 18))
            button.setToolTip(name)
            button.setCursor(Qt.PointingHandCursor)
            button.clicked.connect(lambda checked=False, idx=i: self.switch(idx))
            nav_layout.addWidget(button)
            self.nav_btns.append(button)
            self.nav_names.append(name)
            
        nav_layout.addStretch()

        # Separator line above sidebar footer
        sep2 = QFrame()
        sep2.setFrameShape(QFrame.HLine)
        sep2.setStyleSheet("background-color: #1e2638; border: none; height: 1px; max-height: 1px; margin: 6px 8px;")
        nav_layout.addWidget(sep2)

        # Update button in sidebar footer
        self.update_btn = QPushButton("Check for Updates")
        self.update_btn.setCursor(Qt.PointingHandCursor)
        self.update_btn.setStyleSheet("""
            QPushButton {
                background: #161c28;
                color: #94a3b8;
                font-size: 8pt;
                font-weight: 500;
                padding: 6px;
                border-radius: 6px;
                border: 1px solid #232d3f;
                margin: 2px 8px;
            }
            QPushButton:hover {
                background: #1e293b;
                color: #ffffff;
                border-color: #3b82f6;
            }
        """)
        self.update_btn.clicked.connect(self._on_update_clicked)
        nav_layout.addWidget(self.update_btn)

        # About button in sidebar footer
        self.about_sidebar_btn = QPushButton("About PITC Suite")
        self.about_sidebar_btn.setCursor(Qt.PointingHandCursor)
        self.about_sidebar_btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                color: #64748b;
                font-size: 8pt;
                padding: 5px;
                border: none;
                border-radius: 5px;
                margin: 2px 8px 4px 8px;
            }
            QPushButton:hover {
                background: #18202e;
                color: #cbd5e1;
            }
        """)
        self.about_sidebar_btn.clicked.connect(self._show_about_dialog)
        nav_layout.addWidget(self.about_sidebar_btn)

        self._apply_nav_state()
        outer.addWidget(self.nav)

        # 3. Main Content Area
        content = QWidget()
        content.setObjectName("contentArea")
        content.setStyleSheet("""
            QWidget#contentArea {
                background-color: #0d1117;
            }
        """)
        outer.addWidget(content, 1)
        
        layout = QVBoxLayout(content)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(10)

        # Modern Content Header
        header_bar = QHBoxLayout()
        header_bar.setContentsMargins(2, 2, 2, 4)
        header_bar.setSpacing(10)

        title_col = QVBoxLayout()
        title_col.setSpacing(2)

        self.header_title = QLabel("PITC Report Downloader")
        self.header_title.setStyleSheet("color: #f8fafc; font-size: 15pt; font-weight: bold; letter-spacing: 0.3px;")
        title_col.addWidget(self.header_title)

        self.header_subtitle = QLabel("Automated CDS10 reports, OTP login & filtering")
        self.header_subtitle.setStyleSheet("color: #94a3b8; font-size: 8.5pt;")
        title_col.addWidget(self.header_subtitle)

        header_bar.addLayout(title_col)
        header_bar.addStretch()

        # License status pill in header
        lic = getattr(self, "active_license", None) or load_local_license()
        plan_str = format_plan_type(lic.get("plan_type"))
        is_active = lic.get("status") == "active" or "remaining" in format_expiry_date(lic.get("expires_at"))

        self.license_badge = QPushButton(f"✨ {plan_str}")
        self.license_badge.setCursor(Qt.PointingHandCursor)
        self.license_badge.setToolTip("View license and app details")
        badge_border = "#059669" if is_active else "#2e3a50"
        badge_bg = "#06281e" if is_active else "#161c28"
        badge_fg = "#34d399" if is_active else "#94a3b8"
        self.license_badge.setStyleSheet(f"""
            QPushButton {{
                background: {badge_bg};
                color: {badge_fg};
                border: 1px solid {badge_border};
                border-radius: 12px;
                padding: 4px 12px;
                font-size: 8pt;
                font-weight: 600;
            }}
            QPushButton:hover {{
                background: #1e293b;
                color: #ffffff;
                border-color: #3b82f6;
            }}
        """)
        self.license_badge.clicked.connect(self._show_about_dialog)
        header_bar.addWidget(self.license_badge)

        layout.addLayout(header_bar)

        # Header separator
        header_sep = QFrame()
        header_sep.setFrameShape(QFrame.HLine)
        header_sep.setStyleSheet("background-color: #1e2638; border: none; height: 1px; max-height: 1px;")
        layout.addWidget(header_sep)

        # 4. Stacked Tool Panels
        self.stack = QStackedWidget()
        self.panels = []
        for item in TOOLS:
            name = item[0]
            panel_class = item[1]
            panel = panel_class()
            self.panels.append(panel)
            if panel.layout():
                panel.layout().setContentsMargins(4, 8, 4, 4)
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

        self.switch(0)

        # Check for updates in the background after 1.5 seconds
        QTimer.singleShot(1500, lambda: check_for_updates(self, silent=True))

        # Check if app was just updated to show What's New notes
        self._check_post_update()

    def _init_window_size(self):
        """Restore saved window size or initialize to 700 x 720."""
        settings = QSettings("TNDevelopers", "PITCSuite")
        saved_size = settings.value("window_size")
        saved_geom = settings.value("window_geometry")
        
        if saved_geom and isinstance(saved_geom, (bytes, bytearray, QByteArray)):
            try:
                self.restoreGeometry(saved_geom)
            except Exception:
                pass
                
        if saved_size and isinstance(saved_size, QSize) and saved_size.isValid():
            self.resize(saved_size)
        elif not saved_geom:
            self.resize(700, 720)

    def _save_window_size(self):
        """Persist current window size and geometry so subsequent launches restore them."""
        if not self.isMaximized() and not self.isMinimized() and not self.isFullScreen():
            settings = QSettings("TNDevelopers", "PITCSuite")
            settings.setValue("window_size", self.size())
            settings.setValue("window_geometry", self.saveGeometry())

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._save_window_size()

    def _show_about_dialog(self):
        dlg = AboutDialog(self)
        dlg.exec()

    def switch(self, idx):
        if idx < 0 or idx >= len(TOOLS):
            return

        self.stack.setCurrentIndex(idx)

        # Update top header labels
        tool_info = TOOLS[idx]
        self.header_title.setText(tool_info[0])
        self.header_subtitle.setText(tool_info[3] if len(tool_info) > 3 else "")

        # Smooth moderate page transition animation (fade-in 180ms ease-out)
        curr_widget = self.stack.currentWidget()
        if curr_widget:
            fade_effect = QGraphicsOpacityEffect(curr_widget)
            curr_widget.setGraphicsEffect(fade_effect)
            fade_anim = QPropertyAnimation(fade_effect, b"opacity")
            fade_anim.setDuration(180)
            fade_anim.setStartValue(0.3)
            fade_anim.setEndValue(1.0)
            fade_anim.setEasingCurve(QEasingCurve.OutCubic)
            fade_anim.start(QAbstractAnimation.DeleteWhenStopped)
            self._current_fade_anim = fade_anim

        # Update active nav indicator
        for i, button in enumerate(self.nav_btns):
            button.setProperty("active", "true" if i == idx else "false")
            button.style().unpolish(button)
            button.style().polish(button)

    def toggle_nav(self):
        self.nav_collapsed = not self.nav_collapsed
        
        # Smooth animated sidebar width transition (200ms ease-out cubic)
        target_width = 52 if self.nav_collapsed else 220
        start_width = self.nav.width()
        
        self.nav_anim_group = QParallelAnimationGroup(self)
        
        anim_max = QPropertyAnimation(self.nav, b"maximumWidth")
        anim_max.setDuration(200)
        anim_max.setStartValue(start_width)
        anim_max.setEndValue(target_width)
        anim_max.setEasingCurve(QEasingCurve.OutCubic)
        
        anim_min = QPropertyAnimation(self.nav, b"minimumWidth")
        anim_min.setDuration(200)
        anim_min.setStartValue(start_width)
        anim_min.setEndValue(target_width)
        anim_min.setEasingCurve(QEasingCurve.OutCubic)
        
        self.nav_anim_group.addAnimation(anim_max)
        self.nav_anim_group.addAnimation(anim_min)
        
        self.nav_anim_group.finished.connect(self._apply_nav_state)
        self.nav_anim_group.start(QAbstractAnimation.DeleteWhenStopped)
        
        # Update text labels immediately for responsive feeling
        if self.nav_collapsed:
            self._apply_nav_state()

    def notify_update_available(self, latest_tag=None, notes=None, download_url=None, file_size=0):
        self._has_update = True
        self._latest_tag = latest_tag
        self._latest_notes = notes
        self._download_url = download_url
        self._latest_file_size = file_size
        self.update_btn.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #10b981, stop:1 #059669);
                color: #ffffff;
                font-size: 8pt;
                font-weight: bold;
                padding: 6px;
                border-radius: 6px;
                border: 1px solid #34d399;
                margin: 2px 8px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #34d399, stop:1 #10b981);
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
        target_width = 52 if self.nav_collapsed else 220
        self.nav.setFixedWidth(target_width)
        
        has_update = getattr(self, "_has_update", False)
        if self.nav_collapsed:
            self.nav_brand_lbl.hide()
            self.update_btn.setText("!" if has_update else "🔄")
            self.about_sidebar_btn.setText("ℹ")
        else:
            self.nav_brand_lbl.show()
            self.update_btn.setText("Update Available!" if has_update else "Check for Updates")
            self.about_sidebar_btn.setText("About PITC Suite")
            
        self.update_btn.setToolTip("Check for Updates")
        self.about_sidebar_btn.setToolTip("About PITC Suite")
        
        for button, name in zip(self.nav_btns, self.nav_names):
            if self.nav_collapsed:
                button.setText("")
                button.setStyleSheet(NAV_STYLE + "\nQPushButton { padding: 8px 0; text-align: center; margin: 1px 4px; }")
            else:
                button.setText(name)
                button.setStyleSheet(NAV_STYLE)

    def closeEvent(self, event):
        # Save window size and geometry before closing
        self._save_window_size()
        
        # Gracefully attempt to stop any background workers and drivers
        for panel in self.panels:
            for method in ['cancel', 'halt', 'stop', '_stop']:
                if hasattr(panel, method):
                    try:
                        getattr(panel, method)()
                    except Exception:
                        pass
        super().closeEvent(event)

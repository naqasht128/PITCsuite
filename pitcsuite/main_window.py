import os
from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QLabel, QPushButton, QVBoxLayout, QHBoxLayout,
    QStackedWidget, QFrame, QLineEdit, QSpinBox, QSizePolicy, QScrollArea,
    QAbstractScrollArea, QDialog, QApplication, QGraphicsOpacityEffect,
    QGroupBox, QComboBox, QCheckBox
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
    ("PITC Reports Downloader",       PITCPanel, "pitc_down", "Automated PITC Reports, OTP login, Filtering"),
    ("E-Bills Downloader",            BillSoftPanel, "ebills", "Batch download consumer electricity bills"),
    ("Bill Scrapper",                 BillScraperPanel, "bill_scrap", "Automated web scraper for bill data extraction"),
    ("Payment from 88L",              PaymentExtract88LPanel, "payment_ext", "Extract and parse payment data from 88L reports"),
    ("CP 52 Posting to PDF",          CP52PostingCheckerPanel, "cp52", "Process CP 52 posting records and generate PDF"),
    ("CP 139 Posting/PDF",            CP139PostingCheckerPanel, "cp139", "Verify CP 139 records & generate PDF reports"),
    ("CP 136/137 Posting",            CP136PostingCheckerPanel, "cp136", "Audit and verify CP 136 / 137 posting records"),
    ("CP 22T Posting",                CP22TCheckerPanel, "cp22t", "Posting verification and reconciliation"),
    ("CP21c Converter",               CP21cConverterPanel, "cp21c", "Convert and format CP21c records"),
    ("Assessment Realization",        AssessmentRealizationPanel, "assessment_realization", "Extract government assessment & realization data"),
    ("PDF Merger",                    PDFMerger, "pdf_merger", "Merge, split, and stamp PDF documents"),
    ("File Mover",                    FileMoverPanel, "file_mover", "Organize and move files by batch pattern"),
    ("File Renamer",                  FileRenamerPanel, "file_renamer", "Batch rename files with custom rules"),
    ("Settings",                      SettingsPanel, "settings", "DISCO configuration & login credentials"),
]


class AboutDialog(QDialog):
    """Sleek modal dialog fully matching active application theme palette and typography."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("About")
        self.setFixedSize(500, 520)
        self.setModal(True)
        self.setWindowFlags(Qt.Dialog | Qt.WindowTitleHint | Qt.WindowCloseButtonHint)

        from pitcsuite.ui_helpers import (
            get_current_theme_key, get_current_accent_key, get_current_font_family_key,
            THEMES, ACCENTS, FONTS
        )
        theme_key = get_current_theme_key()
        accent_key = get_current_accent_key()
        palette_data = THEMES.get(theme_key, THEMES["vesper"])
        accent = ACCENTS.get(accent_key, ACCENTS["blue"])
        font_info = FONTS.get(get_current_font_family_key(), FONTS["inter"])
        font_family = font_info["family"]
        is_light = (theme_key == "light")
        brand_heading_color = "#000000" if is_light else palette_data.get("text_main", "#ffffff")

        self.setStyleSheet(f"""
            QDialog {{
                background: {palette_data['window_bg']};
                color: {palette_data['text_sec']};
                font-family: {font_family};
            }}
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 20, 22, 18)
        layout.setSpacing(10)

        # App Header with App Icon (Dark roundish icon)
        header_row = QHBoxLayout()
        header_row.setSpacing(12)

        icon_lbl = QLabel()
        icon_lbl.setPixmap(app_icon().pixmap(44, 44))
        icon_lbl.setFixedSize(44, 44)
        header_row.addWidget(icon_lbl)
        
        title_vbox = QVBoxLayout()
        title_vbox.setSpacing(2)
        
        title_lbl = QLabel("PITC Suite")
        title_lbl.setStyleSheet(f"color: {brand_heading_color}; font-size: 15pt; font-weight: bold; letter-spacing: 0.5px;")
        title_vbox.addWidget(title_lbl)

        ver_lbl = QLabel(f"Version {get_current_version()} · Suite Edition")
        ver_lbl.setStyleSheet(f"color: {palette_data['text_muted']}; font-size: 8.5pt;")
        title_vbox.addWidget(ver_lbl)
        
        header_row.addLayout(title_vbox)
        header_row.addStretch()

        badge = QLabel("TN Developers")
        badge.setStyleSheet(f"background: {palette_data['card_bg']}; border: 1px solid {palette_data['card_border']}; color: {accent['hex']}; font-weight: bold; font-size: 8pt; padding: 4px 10px; border-radius: 4px;")
        header_row.addWidget(badge)
        layout.addLayout(header_row)

        has_update = getattr(parent, "_has_update", False)
        latest_tag = getattr(parent, "_latest_tag", "")
        if has_update and latest_tag:
            banner = QFrame()
            banner.setStyleSheet(f"""
                QFrame {{
                    background: {"#ecfdf5" if is_light else "#064e3b"};
                    border: 1px solid {"#10b981" if is_light else "#059669"};
                    border-radius: 6px;
                }}
            """)
            b_layout = QHBoxLayout(banner)
            b_layout.setContentsMargins(12, 8, 12, 8)
            b_lbl = QLabel(f"✨ <b>Update Available:</b> Version <b>{latest_tag}</b> is ready to install!")
            b_lbl.setStyleSheet(f"color: {"#065f46" if is_light else "#6ee7b7"}; font-size: 8.5pt;")
            b_layout.addWidget(b_lbl)
            layout.addWidget(banner)

        # 1. License Information Card
        lic = getattr(parent, "active_license", None) or load_local_license()
        serial_key = lic.get("serial_key", "Not Activated")
        plan_str = format_plan_type(lic.get("plan_type"))
        expiry_str = format_expiry_date(lic.get("expires_at"))

        lic_card = QFrame()
        lic_card.setStyleSheet(f"""
            QFrame {{
                background: {palette_data['card_bg']};
                border: 1px solid {palette_data['card_border']};
                border-radius: 6px;
            }}
            QLabel {{
                border: none;
                background: transparent;
            }}
        """)
        lic_layout = QVBoxLayout(lic_card)
        lic_layout.setContentsMargins(14, 12, 14, 12)
        lic_layout.setSpacing(6)

        hdr_lbl = QLabel("LICENSE & REGISTRATION")
        hdr_lbl.setStyleSheet(f"color: {palette_data['text_muted']}; font-size: 8pt; font-weight: bold; letter-spacing: 0.5px;")
        lic_layout.addWidget(hdr_lbl)

        key_row = QHBoxLayout()
        key_title = QLabel("Key:")
        key_title.setStyleSheet(f"color: {palette_data['text_sec']}; font-size: 8.5pt; font-weight: 600;")
        key_row.addWidget(key_title)

        key_val = QLabel(serial_key)
        key_val.setStyleSheet(f"color: {accent['hex']}; font-family: 'Cascadia Code', Consolas, monospace; font-size: 8.5pt; font-weight: 600;")
        key_row.addWidget(key_val)
        key_row.addStretch()

        self.change_key_btn = QPushButton("Change Key")
        self.change_key_btn.setCursor(Qt.PointingHandCursor)
        self.change_key_btn.setStyleSheet(f"""
            QPushButton {{
                background: {palette_data['input_bg']};
                color: {palette_data['text_main']};
                font-size: 8pt;
                font-weight: 600;
                padding: 3px 10px;
                border: 1px solid {palette_data['border']};
                border-radius: 4px;
            }}
            QPushButton:hover {{
                background: {palette_data['card_border']};
                border-color: {accent['hex']};
                color: {palette_data['text_main']};
            }}
        """)
        self.change_key_btn.clicked.connect(self._change_serial_key)
        key_row.addWidget(self.change_key_btn)

        self.copy_key_btn = QPushButton("Copy")
        self.copy_key_btn.setCursor(Qt.PointingHandCursor)
        self.copy_key_btn.setStyleSheet(f"""
            QPushButton {{
                background: {palette_data['input_bg']};
                color: {palette_data['text_muted']};
                font-size: 8pt;
                padding: 3px 10px;
                border: 1px solid {palette_data['border']};
                border-radius: 4px;
            }}
            QPushButton:hover {{
                background: {palette_data['card_border']};
                color: {palette_data['text_main']};
            }}
        """)
        self.copy_key_btn.clicked.connect(self._copy_serial_key)
        key_row.addWidget(self.copy_key_btn)
        lic_layout.addLayout(key_row)

        hw_lbl = QLabel(f"Hardware ID:  {get_device_id()}")
        hw_lbl.setStyleSheet(f"color: {palette_data['text_muted']}; font-family: 'Cascadia Code', Consolas, monospace; font-size: 8pt;")
        lic_layout.addWidget(hw_lbl)

        info_row = QHBoxLayout()
        plan_lbl = QLabel(f"Plan: <b style='color:{palette_data['text_main']};'>{plan_str}</b>")
        plan_lbl.setStyleSheet(f"color: {palette_data['text_sec']}; font-size: 8.5pt;")
        info_row.addWidget(plan_lbl)
        info_row.addStretch()

        status_color = "#10b981" if "remaining" in expiry_str or "today" in expiry_str or "Lifetime" in plan_str else palette_data['text_muted']
        status_lbl = QLabel(f"Status: <b style='color:{status_color};'>{expiry_str}</b>")
        status_lbl.setStyleSheet("font-size: 8.5pt;")
        info_row.addWidget(status_lbl)
        lic_layout.addLayout(info_row)

        layout.addWidget(lic_card)

        # 2. Developer Information Card
        dev_card = QFrame()
        dev_card.setStyleSheet(f"""
            QFrame {{
                background: {palette_data['card_bg']};
                border: 1px solid {palette_data['card_border']};
                border-radius: 6px;
            }}
            QLabel {{
                border: none;
                background: transparent;
            }}
        """)
        dev_layout = QVBoxLayout(dev_card)
        dev_layout.setContentsMargins(14, 12, 14, 12)
        dev_layout.setSpacing(5)

        dev_title = QLabel("DEVELOPER SUPPORT")
        dev_title.setStyleSheet(f"color: {palette_data['text_muted']}; font-size: 8pt; font-weight: bold; letter-spacing: 0.5px;")
        dev_layout.addWidget(dev_title)

        org_lbl = QLabel("TN Developers Co.")
        org_lbl.setStyleSheet(f"color: {palette_data['text_main']}; font-size: 9pt; font-weight: 600;")
        dev_layout.addWidget(org_lbl)

        email_lbl = QLabel(f"Email:  <span style='color: {accent['hex']};'>tahir.naqash4@gmail.com</span>")
        email_lbl.setStyleSheet(f"color: {palette_data['text_sec']}; font-size: 8.5pt;")
        dev_layout.addWidget(email_lbl)

        contact_lbl = QLabel("WhatsApp / Contact:  <span style='color: #10b981;'>03339962474</span>")
        contact_lbl.setStyleSheet(f"color: {palette_data['text_sec']}; font-size: 8.5pt;")
        dev_layout.addWidget(contact_lbl)

        layout.addWidget(dev_card)
        layout.addStretch()

        # Bottom Button Row
        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)

        whats_new_btn = QPushButton("What's New")
        whats_new_btn.setCursor(Qt.PointingHandCursor)
        whats_new_btn.setStyleSheet(f"""
            QPushButton {{
                background: {palette_data['input_bg']};
                color: {palette_data['text_main']};
                font-size: 8.5pt;
                font-weight: 600;
                padding: 6px 14px;
                border-radius: 4px;
                border: 1px solid {palette_data['border']};
            }}
            QPushButton:hover {{
                background: {palette_data['card_border']};
                border-color: {accent['hex']};
            }}
        """)
        whats_new_btn.clicked.connect(self._show_whats_new)
        btn_row.addWidget(whats_new_btn)

        # Check for Updates / Update Available button (Request 9, 10)
        self.update_action_btn = QPushButton()
        self.update_action_btn.setCursor(Qt.PointingHandCursor)
        if has_update:
            self.update_action_btn.setText(f"Update Available ({latest_tag})" if latest_tag else "Update Available")
            self.update_action_btn.setStyleSheet("""
                QPushButton {
                    background: #10b981;
                    color: #ffffff;
                    font-size: 8.5pt;
                    font-weight: 600;
                    padding: 6px 14px;
                    border-radius: 4px;
                    border: none;
                }
                QPushButton:hover {
                    background: #059669;
                }
                QPushButton:pressed {
                    background: #047857;
                }
            """)
            self.update_action_btn.clicked.connect(self._on_update_available_clicked)
        else:
            self.update_action_btn.setText("Check for Updates")
            self.update_action_btn.setStyleSheet(f"""
                QPushButton {{
                    background: {palette_data['input_bg']};
                    color: {palette_data['text_main']};
                    font-size: 8.5pt;
                    font-weight: 600;
                    padding: 6px 14px;
                    border-radius: 4px;
                    border: 1px solid {palette_data['border']};
                }}
                QPushButton:hover {{
                    background: {palette_data['card_border']};
                    border-color: {palette_data['border_hover']};
                }}
            """)
            self.update_action_btn.clicked.connect(self._run_check_updates)
        btn_row.addWidget(self.update_action_btn)

        btn_row.addStretch()

        close_btn = QPushButton("Close")
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.setStyleSheet(f"""
            QPushButton {{
                background: {palette_data['input_bg']};
                color: {palette_data['text_main']};
                font-size: 8.5pt;
                padding: 6px 18px;
                border-radius: 4px;
                border: 1px solid {palette_data['border']};
            }}
            QPushButton:hover {{
                background: {palette_data['card_border']};
                color: {palette_data['text_main']};
            }}
        """)
        close_btn.clicked.connect(self.accept)
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

    def _on_update_available_clicked(self):
        parent = self.parent()
        self.accept()
        if parent and hasattr(parent, '_on_update_clicked'):
            parent._on_update_clicked()

    def _run_check_updates(self):
        parent = self.parent()
        self.update_action_btn.setText("Checking...")
        self.update_action_btn.setEnabled(False)
        self.accept()
        if parent:
            check_for_updates(parent, silent=False)

    def _copy_serial_key(self):
        lic = getattr(self.parent(), "active_license", None) or load_local_license()
        key = lic.get("serial_key", "")
        if key:
            QApplication.clipboard().setText(key)
            self.copy_key_btn.setText("Copied!")
            self.copy_key_btn.setStyleSheet("""
                QPushButton {
                    background: #142e23;
                    color: #89d185;
                    font-size: 8pt;
                    padding: 3px 8px;
                    border: 1px solid #2ea043;
                    border-radius: 4px;
                }
            """)
            QTimer.singleShot(2000, lambda: self._reset_copy_btn() if not self.isHidden() else None)

    def _reset_copy_btn(self):
        from pitcsuite.ui_helpers import get_current_theme_key, THEMES
        palette_data = THEMES.get(get_current_theme_key(), THEMES["vesper"])
        self.copy_key_btn.setText("Copy")
        self.copy_key_btn.setStyleSheet(f"""
            QPushButton {{
                background: {palette_data['input_bg']};
                color: {palette_data['text_muted']};
                font-size: 8pt;
                padding: 3px 8px;
                border: 1px solid {palette_data['border']};
                border-radius: 4px;
            }}
            QPushButton:hover {{
                background: {palette_data['card_border']};
                color: {palette_data['text_main']};
            }}
        """)

    def _change_serial_key(self):
        parent = self.parent()
        self.accept()
        gate = LicenseGateDialog(parent)
        gate.exec()

    def _show_whats_new(self):
        notes = get_release_notes(get_current_version())
        dlg = UpToDateDialog(get_current_version(), notes, self, title="What's New in PITC Suite")
        dlg.exec()

class FramelessScrollArea(QScrollArea):
    """
    Scroll area with smooth vertical scrolling and middle-click pan.
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWidgetResizable(True)
        self.setFrameShape(QFrame.NoFrame)
        self.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self._is_panning = False
        self._pan_start = None
        self._h_start = 0
        self._v_start = 0

    def wheelEvent(self, event):
        super().wheelEvent(event)

    def mousePressEvent(self, event):
        if event.button() == Qt.MiddleButton:
            self._is_panning = True
            pos = event.position().toPoint() if hasattr(event, "position") else event.pos()
            self._pan_start = pos
            self._v_start = self.verticalScrollBar().value()
            self.viewport().setCursor(Qt.ClosedHandCursor)
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._is_panning and self._pan_start is not None:
            pos = event.position().toPoint() if hasattr(event, "position") else event.pos()
            delta = pos - self._pan_start
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
    def __init__(self, scroll_area: FramelessScrollArea):
        super().__init__(scroll_area)
        self.scroll_area = scroll_area

    def eventFilter(self, watched, event):
        etype = event.type()
        if etype == QEvent.MouseButtonPress and event.button() == Qt.MiddleButton:
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
        self.setMinimumSize(540, 480)
        
        # 1. Window size restoration (800 x 700 initial, or last saved)
        self._init_window_size()

        root = QWidget()
        root.setObjectName("centralWidget")
        self.setCentralWidget(root)
        outer = QHBoxLayout(root)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # 2. Antigravity Exact Activity Bar
        from pitcsuite.ui_helpers import get_current_theme_key, THEMES
        theme_key = get_current_theme_key()
        palette_data = THEMES.get(theme_key, THEMES["vesper"])
        is_light = (theme_key == "light")
        brand_color = "#000000" if is_light else palette_data.get("text_main", "#ffffff")

        settings = QSettings("TNDevelopers", "PITCSuite")
        sidebar_pref = str(settings.value("sidebar_style", "expanded")).lower()
        self.nav_collapsed = (sidebar_pref == "compact")
        self.nav = QWidget()
        self.nav.setObjectName("navSidebar")
        self.nav.setStyleSheet(f"""
            QWidget#navSidebar {{
                background-color: {palette_data['nav_bg']};
                border-right: 1px solid {palette_data['nav_border']};
            }}
        """)
        nav_layout = QVBoxLayout(self.nav)
        nav_layout.setContentsMargins(0, 4, 0, 4)
        nav_layout.setSpacing(2)

        # Activity Bar Header: Toggle button & Branding
        nav_header = QHBoxLayout()
        nav_header.setContentsMargins(6, 2, 6, 4)
        nav_header.setSpacing(6)

        self.nav_menu_btn = QPushButton("☰")
        self.nav_menu_btn.setFixedSize(36, 36)
        self.nav_menu_btn.setToolTip("Toggle activity labels")
        self.nav_menu_btn.setCursor(Qt.PointingHandCursor)
        self.nav_menu_btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                color: #858585;
                border: none;
                border-radius: 4px;
                font-size: 11pt;
                padding: 0;
            }
            QPushButton:hover {
                background: #2a2d2e;
                color: #ffffff;
            }
        """)
        self.nav_menu_btn.clicked.connect(self.toggle_nav)
        nav_header.addWidget(self.nav_menu_btn)

        self.nav_brand_lbl = QLabel("PITC Suite")
        self.nav_brand_lbl.setStyleSheet(f"color: {brand_color}; font-size: 10.5pt; font-weight: bold; letter-spacing: 0.5px;")
        nav_header.addWidget(self.nav_brand_lbl)
        nav_header.addStretch()
        nav_layout.addLayout(nav_header)

        sep1 = QFrame()
        sep1.setFrameShape(QFrame.HLine)
        sep1.setStyleSheet("background-color: #2b2b2b; border: none; height: 1px; max-height: 1px; margin: 0 4px 4px 4px;")
        nav_layout.addWidget(sep1)

        # Tool Activity Bar Buttons (Exact Antigravity Style)
        self.nav_btns = []
        self.nav_names = []
        self.nav_icons = []
        for i, item in enumerate(TOOLS):
            name = item[0]
            icon_name = item[2]
            button = QPushButton(name)
            button.setIcon(get_icon(icon_name, is_light=is_light))
            button.setIconSize(QSize(20, 20))
            button.setToolTip(name)
            button.setCursor(Qt.PointingHandCursor)
            button.setFixedHeight(38)
            button.clicked.connect(lambda checked=False, idx=i: self.switch(idx))
            nav_layout.addWidget(button)
            self.nav_btns.append(button)
            self.nav_names.append(name)
            self.nav_icons.append(icon_name)
            
        # About Button (directly continuous with above tools and Settings)
        self.about_sidebar_btn = QPushButton("About")
        self.about_sidebar_btn.setIcon(get_icon("info", is_light=is_light))
        self.about_sidebar_btn.setIconSize(QSize(20, 20))
        self.about_sidebar_btn.setCursor(Qt.PointingHandCursor)
        self.about_sidebar_btn.setFixedHeight(38)
        self.about_sidebar_btn.setToolTip("About")
        self.about_sidebar_btn.clicked.connect(self._show_about_dialog)
        nav_layout.addWidget(self.about_sidebar_btn)

        nav_layout.addStretch()

        self._apply_nav_state()
        outer.addWidget(self.nav)

        # 3. Main Content Area (Antigravity Exact Canvas #1e1e1e)
        content = QWidget()
        content.setObjectName("contentArea")
        content.setStyleSheet("""
            QWidget#contentArea {
                background-color: #1e1e1e;
            }
        """)
        outer.addWidget(content, 1)
        
        self.content_layout = QVBoxLayout(content)
        self.content_layout.setContentsMargins(14, 10, 14, 10)
        self.content_layout.setSpacing(8)

        # Content Header Bar (Antigravity Style)
        header_bar = QHBoxLayout()
        header_bar.setContentsMargins(2, 2, 2, 2)
        header_bar.setSpacing(8)

        title_col = QVBoxLayout()
        title_col.setSpacing(1)

        heading_color = "#000000" if is_light else palette_data.get("text_main", "#ffffff")
        self.header_title = QLabel("PITC Reports Downloader")
        self.header_title.setStyleSheet(f"color: {heading_color}; font-size: 13pt; font-weight: 600; letter-spacing: 0.3px;")
        title_col.addWidget(self.header_title)

        self.header_subtitle = QLabel("Automated PITC Reports, OTP login, Filtering")
        self.header_subtitle.setStyleSheet(f"color: {palette_data.get('text_muted', '#9d9d9d')}; font-size: 8.5pt;")
        title_col.addWidget(self.header_subtitle)

        header_bar.addLayout(title_col)
        header_bar.addStretch()

        self.content_layout.addLayout(header_bar)

        header_sep = QFrame()
        header_sep.setFrameShape(QFrame.HLine)
        header_sep.setStyleSheet("background-color: #2b2b2b; border: none; height: 1px; max-height: 1px;")
        self.content_layout.addWidget(header_sep)

        # 4. Stacked Tool Panels
        self.stack = QStackedWidget()
        self.panels = []
        for item in TOOLS:
            name = item[0]
            panel_class = item[1]
            panel = panel_class()
            self.panels.append(panel)
            
            # Clean up ONLY duplicate internal titles inside panels so top header is clean
            if panel.layout():
                panel.layout().setContentsMargins(6, 6, 6, 6)
                panel.layout().setAlignment(Qt.AlignTop)
                for idx in range(min(3, panel.layout().count())):
                    child_item = panel.layout().itemAt(idx)
                    if not child_item or not child_item.widget():
                        continue
                    w = child_item.widget()
                    if isinstance(w, QGroupBox):
                        break
                    if isinstance(w, QLabel):
                        txt = w.text().strip()
                        # Never hide input labels like 'Excel file:', 'Username:', 'Bills URL:'
                        if txt.endswith(":") or len(txt) < 3:
                            break
                        if (name.lower() in txt.lower() or 
                            txt.lower() in name.lower() or 
                            "converter" in txt.lower() or "extractor" in txt.lower() or 
                            "checker" in txt.lower() or "downloader" in txt.lower() or 
                            "merger" in txt.lower() or "mover" in txt.lower() or "renamer" in txt.lower() or
                            "settings" in txt.lower() or "scrapper" in txt.lower() or "posting" in txt.lower()):
                            w.hide()
                        elif w.font().bold() and len(txt) < 45:
                            w.hide()
                        else:
                            break
                    elif isinstance(w, QFrame) and w.frameShape() == QFrame.HLine:
                        w.hide()
                    else:
                        break

            # Word-wrap all labels so long texts never inflate panel minimum width
            for qlbl in panel.findChildren(QLabel):
                qlbl.setWordWrap(True)

            # Allow comboboxes to naturally adjust and show their dropdown arrow
            for qcb in panel.findChildren(QComboBox):
                qcb.setSizeAdjustPolicy(QComboBox.AdjustToContents)
                qcb.setMinimumContentsLength(8)
                if qcb.view():
                    qcb.view().setTextElideMode(Qt.ElideMiddle)
                            
            for button in panel.findChildren(QPushButton):
                if button.property("role") is None and not button.styleSheet():
                    button.setProperty("role", "secondary")
                button.setMinimumWidth(0)
                button.setMaximumWidth(16777215)
                
            for edit in panel.findChildren(QLineEdit):
                if isinstance(edit.parentWidget(), QSpinBox):
                    continue
                edit.setMinimumWidth(0)
                edit.setMaximumWidth(16777215)
                edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

            for chk in panel.findChildren(QCheckBox):
                chk.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
                
            panel.setMinimumWidth(0)
            scroll = FramelessScrollArea()
            scroll.setWidgetResizable(True)
            scroll.setWidget(panel)
            filt = ScrollEventFilter(scroll)
            panel.installEventFilter(filt)
            for child in panel.findChildren(QWidget):
                child.installEventFilter(filt)
                if isinstance(child, QAbstractScrollArea):
                    child.viewport().installEventFilter(filt)
                    
            for area in [scroll, *panel.findChildren(QAbstractScrollArea)]:
                area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
                area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
                
            self.stack.addWidget(scroll)
            
        self.content_layout.addWidget(self.stack, 1)

        self.switch(0)

        # Check for updates in the background after 1.5 seconds if enabled
        settings = QSettings("TNDevelopers", "PITCSuite")
        if settings.value("auto_update_check", True, type=bool):
            QTimer.singleShot(1500, lambda: check_for_updates(self, silent=True))

        # Check if app was just updated to show What's New notes
        self._check_post_update()

    def _init_window_size(self):
        """Restore saved window size or initialize according to launch preference (800 x 650 default)."""
        settings = QSettings("TNDevelopers", "PITCSuite")
        launch_mode = str(settings.value("launch_mode", "saved")).lower()

        if launch_mode == "maximized":
            self.showMaximized()
            return

        if launch_mode == "saved":
            saved_geom = settings.value("window_geometry")
            if saved_geom and isinstance(saved_geom, (bytes, bytearray, QByteArray)):
                try:
                    self.restoreGeometry(saved_geom)
                except Exception:
                    pass
            saved_size = settings.value("window_size")
            if saved_size and isinstance(saved_size, QSize) and saved_size.isValid():
                self.resize(saved_size)
                return

        # Default size: 800 x 650
        self.resize(800, 650)

    def _save_window_size(self):
        """Persist current window size and geometry so subsequent launches restore them."""
        if not self.isMaximized() and not self.isMinimized() and not self.isFullScreen():
            settings = QSettings("TNDevelopers", "PITCSuite")
            settings.setValue("window_size", self.size())
            settings.setValue("window_geometry", self.saveGeometry())

    def set_sidebar_collapsed(self, collapsed: bool):
        if self.nav_collapsed != collapsed:
            self.toggle_nav()

    def _on_theme_updated(self, theme_key, accent_key, font_size_key, density_key, font_family_key=None):
        self._current_theme_key = theme_key
        self._current_accent_key = accent_key
        from pitcsuite.ui_helpers import THEMES, ACCENTS, NAV_STYLE
        palette_data = THEMES.get(theme_key, THEMES["vesper"])
        accent = ACCENTS.get(accent_key, ACCENTS["red"])
        
        self.nav.setStyleSheet(f"""
            QWidget#navSidebar {{
                background-color: {palette_data['nav_bg']};
                border-right: 1px solid {palette_data['nav_border']};
            }}
        """)
        
        content = self.centralWidget().findChild(QWidget, "contentArea")
        if content:
            content.setStyleSheet(f"""
                QWidget#contentArea {{
                    background-color: {palette_data['window_bg']};
                }}
            """)

        is_light = (theme_key == "light")
        brand_color = "#000000" if is_light else palette_data.get("text_main", "#ffffff")
        self.nav_brand_lbl.setStyleSheet(f"color: {brand_color}; font-size: 10.5pt; font-weight: bold; letter-spacing: 0.5px;")

        heading_color = "#000000" if is_light else palette_data.get("text_main", "#ffffff")
        self.header_title.setStyleSheet(f"color: {heading_color}; font-size: 13pt; font-weight: 600; letter-spacing: 0.3px;")
        self.header_subtitle.setStyleSheet(f"color: {palette_data['text_muted']}; font-size: 8.5pt;")

        # Update status line styling across all tool panels (black on light UI, white on dark UI)
        from pitcsuite.ui_helpers import get_status_style
        status_style = get_status_style(theme_key)
        for panel in getattr(self, "panels", []):
            if hasattr(panel, "_update_accent_buttons_style"):
                panel._update_accent_buttons_style(theme_key)
        for panel in getattr(self, "panels", []):
            for lbl_w in panel.findChildren(QLabel):
                if lbl_w.objectName() == "statusLabel" or lbl_w.property("role") == "status":
                    lbl_w.setStyleSheet(status_style)

        self._apply_nav_state()

    def _adapt_responsive_ui(self, width: int, height: int):
        """Dynamically adapt fonts, margins, and layout to window adjustments."""
        from pitcsuite.ui_helpers import THEMES, get_current_theme_key
        theme_key = get_current_theme_key()
        palette_data = THEMES.get(theme_key, THEMES["vesper"])
        text_main = "#000000" if theme_key == "light" else palette_data.get("text_main", "#ffffff")
        text_muted = palette_data.get("text_muted", "#858585")

        if width < 700 or height < 580:
            self.header_title.setStyleSheet(f"color: {text_main}; font-size: 11pt; font-weight: 600; letter-spacing: 0.2px;")
            self.header_subtitle.setStyleSheet(f"color: {text_muted}; font-size: 7.5pt;")
            self.content_layout.setContentsMargins(6, 4, 6, 4)
            self.content_layout.setSpacing(5)
            if not self.nav_collapsed and width < 650:
                self.toggle_nav()
        elif width < 820:
            self.header_title.setStyleSheet(f"color: {text_main}; font-size: 12pt; font-weight: 600; letter-spacing: 0.3px;")
            self.header_subtitle.setStyleSheet(f"color: {text_muted}; font-size: 8pt;")
            self.content_layout.setContentsMargins(10, 6, 10, 6)
            self.content_layout.setSpacing(6)
        else:
            self.header_title.setStyleSheet(f"color: {text_main}; font-size: 13pt; font-weight: 600; letter-spacing: 0.3px;")
            self.header_subtitle.setStyleSheet(f"color: {text_muted}; font-size: 8.5pt;")
            self.content_layout.setContentsMargins(14, 8, 14, 8)
            self.content_layout.setSpacing(8)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._save_window_size()
        self._adapt_responsive_ui(self.width(), self.height())

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

        # Smooth moderate page transition animation (fade-in 180ms ease-out) if enabled
        settings = QSettings("TNDevelopers", "PITCSuite")
        smooth_anim = settings.value("smooth_anim", True, type=bool)
        if smooth_anim:
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

        # Update active nav indicator (Antigravity left white border)
        for i, button in enumerate(self.nav_btns):
            button.setProperty("active", "true" if i == idx else "false")
            button.style().unpolish(button)
            button.style().polish(button)

    def toggle_nav(self):
        self.nav_collapsed = not self.nav_collapsed
        settings = QSettings("TNDevelopers", "PITCSuite")
        settings.setValue("sidebar_style", "compact" if self.nav_collapsed else "expanded")
        
        target_width = 48 if self.nav_collapsed else 180
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
        
        if self.nav_collapsed:
            self._apply_nav_state()

    def notify_update_available(self, latest_tag=None, notes=None, download_url=None, file_size=0):
        self._has_update = True
        self._latest_tag = latest_tag
        self._latest_notes = notes
        self._download_url = download_url
        self._latest_file_size = file_size
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
        target_width = 48 if self.nav_collapsed else 180
        self.nav.setFixedWidth(target_width)
        
        from pitcsuite.ui_helpers import get_nav_style, get_current_theme_key, THEMES
        theme_key = getattr(self, "_current_theme_key", None) or get_current_theme_key()
        palette_data = THEMES.get(theme_key, THEMES["vesper"])
        is_light = (theme_key == "light")
        brand_color = "#000000" if is_light else palette_data.get("text_main", "#ffffff")

        has_update = getattr(self, "_has_update", False)
        accent_key = getattr(self, "_current_accent_key", None)
        current_nav_style = get_nav_style(accent_key=accent_key, theme_key=theme_key)
        collapsed_btn_style = current_nav_style + "\nQPushButton { padding: 8px 0; text-align: center; margin: 1px 0px; }"

        # Hamburger toggle button styling
        self.nav_menu_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                color: {"#1e293b" if is_light else "#858585"};
                border: none;
                border-radius: 4px;
                font-size: 11pt;
                padding: 0;
            }}
            QPushButton:hover {{
                background: {palette_data['item_hover']};
                color: {palette_data['text_main']};
            }}
        """)

        # Footer Button (About / Update Available in Emerald Green)
        if has_update:
            self.about_sidebar_btn.setIcon(get_icon("update_available", color="#ffffff"))
            self.about_sidebar_btn.setIconSize(QSize(20, 20))
            emerald_style = """
                QPushButton {
                    background-color: #10b981;
                    color: #ffffff;
                    border: none;
                    border-radius: 4px;
                    padding: 6px 10px;
                    text-align: left;
                    font-size: 8.5pt;
                    font-weight: 600;
                    margin: 2px 6px;
                }
                QPushButton:hover {
                    background-color: #059669;
                    color: #ffffff;
                }
                QPushButton:pressed {
                    background-color: #047857;
                }
            """
            collapsed_emerald_style = """
                QPushButton {
                    background-color: #10b981;
                    color: #ffffff;
                    border: none;
                    border-radius: 4px;
                    padding: 6px 0;
                    text-align: center;
                    margin: 2px 4px;
                }
                QPushButton:hover {
                    background-color: #059669;
                }
                QPushButton:pressed {
                    background-color: #047857;
                }
            """
            if self.nav_collapsed:
                self.about_sidebar_btn.setText("")
                self.about_sidebar_btn.setStyleSheet(collapsed_emerald_style)
            else:
                self.about_sidebar_btn.setText("Update Available")
                self.about_sidebar_btn.setStyleSheet(emerald_style)
            self.about_sidebar_btn.setToolTip("Update Available - Click to view")
        else:
            self.about_sidebar_btn.setIcon(get_icon("info", is_light=is_light))
            self.about_sidebar_btn.setIconSize(QSize(20, 20))
            if self.nav_collapsed:
                self.about_sidebar_btn.setText("")
                self.about_sidebar_btn.setStyleSheet(collapsed_btn_style)
            else:
                self.about_sidebar_btn.setText("About")
                self.about_sidebar_btn.setStyleSheet(current_nav_style)
            self.about_sidebar_btn.setToolTip("About")

        if self.nav_collapsed:
            self.nav_brand_lbl.hide()
        else:
            self.nav_brand_lbl.setStyleSheet(f"color: {brand_color}; font-size: 10.5pt; font-weight: bold; letter-spacing: 0.5px;")
            self.nav_brand_lbl.show()

        # Update Tool Buttons & Icons for crisp light/dark contrast
        nav_icons = getattr(self, "nav_icons", None)
        for i, (button, name) in enumerate(zip(self.nav_btns, self.nav_names)):
            if nav_icons and i < len(nav_icons):
                button.setIcon(get_icon(nav_icons[i], is_light=is_light))
            if self.nav_collapsed:
                button.setText("")
                button.setStyleSheet(collapsed_btn_style)
            else:
                button.setText(name)
                button.setStyleSheet(current_nav_style)

    def closeEvent(self, event):
        self._save_window_size()
        settings = QSettings("TNDevelopers", "PITCSuite")
        settings.setValue("sidebar_style", "compact" if self.nav_collapsed else "expanded")
        for panel in self.panels:
            for method in ['cancel', 'halt', 'stop', '_stop']:
                if hasattr(panel, method):
                    try:
                        getattr(panel, method)()
                    except Exception:
                        pass
        super().closeEvent(event)

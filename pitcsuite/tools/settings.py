import os
import re
import configparser
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton,
    QLineEdit, QFileDialog, QCheckBox, QTextEdit, QProgressBar,
    QVBoxLayout, QHBoxLayout, QGroupBox, QMessageBox, QComboBox,
    QSpinBox, QTabWidget, QRadioButton, QButtonGroup, QGridLayout,
    QFrame
)
from PySide6.QtCore import Signal, QObject, QThread, QSettings, Qt, QSize
from PySide6.QtGui import QColor, QCursor

from pitcsuite.ui_helpers import (
    lbl, hline, START_BTN, STOP_BTN, THEMES, ACCENTS, FONT_SIZES,
    apply_theme, get_current_theme_key, get_current_accent_key,
    get_current_font_size_key, get_current_density_key
)
from pitcsuite.config import config_path as _config_path
from pitcsuite.auto_updater import get_current_version
from pitcsuite.license_manager import load_local_license, get_device_id, format_plan_type, format_expiry_date


class SettingsPanel(QWidget):
    DISCOS = [
        "PESCO", "LESCO", "GEPCO", "FESCO", "IESCO", "MEPCO",
        "HAZECO", "SEPCO", "HESCO", "QESCO", "TESCO"
    ]

    def __init__(self):
        super().__init__()
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(6, 6, 6, 6)
        main_layout.setSpacing(10)

        # Tab Widget for Antigravity categorized settings
        self.tabs = QTabWidget()
        
        # Tab 1: Personalization
        self.tab_personalization = self._create_personalization_tab()
        self.tabs.addTab(self.tab_personalization, "🎨  Personalization")

        # Tab 2: PITC Credentials
        self.tab_credentials = self._create_credentials_tab()
        self.tabs.addTab(self.tab_credentials, "🔑  PITC Credentials")

        # Tab 3: System & Defaults
        self.tab_system = self._create_system_tab()
        self.tabs.addTab(self.tab_system, "⚙️  System and Paths")

        main_layout.addWidget(self.tabs, 1)

        # Load all existing settings
        self.load_all()

    # =================================================================
    # TAB 1: PERSONALIZATION (Antigravity Xerox)
    # =================================================================
    def _create_personalization_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(12)

        # 1. Appearance / Theme Mode
        theme_box = QGroupBox("Theme Mode")
        tb_layout = QGridLayout(theme_box)
        tb_layout.setHorizontalSpacing(10)
        tb_layout.setVerticalSpacing(8)

        tb_layout.addWidget(QLabel("Color Theme:"), 0, 0)
        self.theme_cb = QComboBox()
        self.theme_cb.addItem("Vesper (Default)", "vesper")
        self.theme_cb.addItem("Slate (Cobalt Dark)", "slate")
        self.theme_cb.addItem("Dark", "dark")
        self.theme_cb.addItem("Midnight (OLED Pitch Black)", "midnight")
        self.theme_cb.addItem("Clean Light Modern", "light")
        self.theme_cb.currentIndexChanged.connect(self._on_live_preview_changed)
        tb_layout.addWidget(self.theme_cb, 0, 1)

        layout.addWidget(theme_box)

        # 2. Accent Color Palette
        accent_box = QGroupBox("Accent Brand Color")
        ab_layout = QVBoxLayout(accent_box)
        ab_layout.setSpacing(8)

        color_grid = QGridLayout()
        color_grid.setHorizontalSpacing(8)
        color_grid.setVerticalSpacing(8)
        self.accent_btns = {}
        self.accent_group = QButtonGroup(self)
        for idx, (key, info) in enumerate(ACCENTS.items()):
            btn = QPushButton(f"● {info['name']}")
            btn.setCheckable(True)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setProperty("accent_key", key)
            btn.clicked.connect(lambda checked=False, k=key: self._select_accent(k))
            self.accent_group.addButton(btn)
            self.accent_btns[key] = btn
            color_grid.addWidget(btn, idx // 3, idx % 3)

        self._update_accent_buttons_style()

        ab_layout.addLayout(color_grid)
        layout.addWidget(accent_box)

        # 3. Typography, UI Scale & Density
        scale_box = QGroupBox("Interface Scaling, Typography and Density")
        sb_layout = QGridLayout(scale_box)
        sb_layout.setHorizontalSpacing(10)
        sb_layout.setVerticalSpacing(8)

        sb_layout.addWidget(QLabel("Font Family:"), 0, 0)
        self.font_cb = QComboBox()
        self.font_cb.addItem("Inter", "inter")
        self.font_cb.addItem("System UI", "system")
        self.font_cb.addItem("Google Sans", "google_sans")
        self.font_cb.currentIndexChanged.connect(self._on_live_preview_changed)
        sb_layout.addWidget(self.font_cb, 0, 1)

        sb_layout.addWidget(QLabel("UI Scale:"), 1, 0)
        self.scale_cb = QComboBox()
        self.scale_cb.addItem("Compact (8.0pt / 90% Scale)", "compact")
        self.scale_cb.addItem("Standard (8.5pt / 100% Scale)", "standard")
        self.scale_cb.addItem("Large (9.5pt / 110% Scale)", "large")
        self.scale_cb.addItem("Extra Large (10.5pt / 125% Scale)", "xlarge")
        self.scale_cb.currentIndexChanged.connect(self._on_live_preview_changed)
        sb_layout.addWidget(self.scale_cb, 1, 1)

        sb_layout.addWidget(QLabel("Density:"), 2, 0)
        self.density_cb = QComboBox()
        self.density_cb.addItem("Compact (Tighter layout for small screens)", "compact")
        self.density_cb.addItem("Standard (Balanced spacing)", "standard")
        self.density_cb.currentIndexChanged.connect(self._on_live_preview_changed)
        sb_layout.addWidget(self.density_cb, 2, 1)

        layout.addWidget(scale_box)

        # 4. Window & Navigation Behavior
        win_box = QGroupBox("Navigation and Window Behavior")
        wb_layout = QGridLayout(win_box)
        wb_layout.setHorizontalSpacing(10)
        wb_layout.setVerticalSpacing(8)

        wb_layout.addWidget(QLabel("Default Activity Bar:"), 0, 0)
        self.sidebar_style_cb = QComboBox()
        self.sidebar_style_cb.addItem("Expanded Rail (180px with Labels)", "expanded")
        self.sidebar_style_cb.addItem("Compact Rail (48px Icons Only)", "compact")
        wb_layout.addWidget(self.sidebar_style_cb, 0, 1)

        wb_layout.addWidget(QLabel("Launch Window Size:"), 1, 0)
        self.launch_mode_cb = QComboBox()
        self.launch_mode_cb.addItem("Remember Last Window Geometry", "saved")
        self.launch_mode_cb.addItem("Standard Default (800 x 650)", "default")
        self.launch_mode_cb.addItem("Maximized Window", "maximized")
        wb_layout.addWidget(self.launch_mode_cb, 1, 1)

        self.chk_smooth_anim = QCheckBox("Smooth page fade-in animations on tool switch")
        self.chk_smooth_anim.setChecked(True)
        wb_layout.addWidget(self.chk_smooth_anim, 2, 0, 1, 2)

        self.chk_auto_update = QCheckBox("Automatically check for updates on startup")
        self.chk_auto_update.setChecked(True)
        wb_layout.addWidget(self.chk_auto_update, 3, 0, 1, 2)

        layout.addWidget(win_box)

        # Bottom Buttons
        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)

        self.btn_save_personalization = QPushButton("Save and Apply Personalization")
        self.btn_save_personalization.setProperty("role", "primary")
        self.btn_save_personalization.clicked.connect(self.save_personalization)
        btn_row.addWidget(self.btn_save_personalization)

        self.btn_reset_personalization = QPushButton("Reset to Defaults")
        self.btn_reset_personalization.clicked.connect(self.reset_personalization)
        btn_row.addWidget(self.btn_reset_personalization)

        btn_row.addStretch()
        layout.addLayout(btn_row)

        self.personalization_status = QLabel("")
        self.personalization_status.setStyleSheet("color:#4c4; font-size:11px;")
        layout.addWidget(self.personalization_status)

        layout.addStretch()
        return w

    # =================================================================
    # TAB 2: PITC CREDENTIALS
    # =================================================================
    def _create_credentials_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(12)

        g = QGroupBox("PITC Login (cp.pitc.com.pk)")
        gv = QGridLayout(g)
        gv.setHorizontalSpacing(10)
        gv.setVerticalSpacing(10)

        # Row 0: DISCO
        gv.addWidget(QLabel("Default DISCO:"), 0, 0)
        self.disco_cb = QComboBox()
        self.disco_cb.addItems(self.DISCOS)
        self.disco_cb.setCurrentText("PESCO")
        gv.addWidget(self.disco_cb, 0, 1, 1, 2)

        # Row 1: Username
        gv.addWidget(QLabel("Username:"), 1, 0)
        self.user_ed = QLineEdit()
        self.user_ed.setPlaceholderText("Enter your PITC username (e.g. 26IPCEO)")
        gv.addWidget(self.user_ed, 1, 1, 1, 2)

        # Row 2: Password
        gv.addWidget(QLabel("Password:"), 2, 0)
        self.pass_ed = QLineEdit()
        self.pass_ed.setEchoMode(QLineEdit.Password)
        self.pass_ed.setPlaceholderText("Enter your PITC password")
        gv.addWidget(self.pass_ed, 2, 1)

        self.chk_show = QCheckBox("Show password")
        self.chk_show.toggled.connect(lambda on: self.pass_ed.setEchoMode(QLineEdit.Normal if on else QLineEdit.Password))
        gv.addWidget(self.chk_show, 2, 2)

        # Row 3: OTP Secret
        gv.addWidget(QLabel("OTP Secret:"), 3, 0)
        self.otp_ed = QLineEdit()
        self.otp_ed.setEchoMode(QLineEdit.Password)
        self.otp_ed.setPlaceholderText("TOTP base32 secret (leave blank if no OTP required)")
        gv.addWidget(self.otp_ed, 3, 1)

        self.chk_otp_show = QCheckBox("Show")
        self.chk_otp_show.toggled.connect(lambda on: self.otp_ed.setEchoMode(QLineEdit.Normal if on else QLineEdit.Password))
        gv.addWidget(self.chk_otp_show, 3, 2)

        gv.setColumnStretch(1, 1)
        layout.addWidget(g)

        self.path_lbl = QLabel(f"Saved to:  {_config_path()}")
        self.path_lbl.setStyleSheet("color:#858585; font-size:10px;")
        self.path_lbl.setWordWrap(True)
        self.path_lbl.setMinimumWidth(0)
        layout.addWidget(self.path_lbl)

        bh = QHBoxLayout()
        self.btn_save_cred = QPushButton("Save Credentials")
        self.btn_save_cred.setProperty("role", "primary")
        self.btn_save_cred.clicked.connect(self.save_credentials)
        bh.addWidget(self.btn_save_cred)
        bh.addStretch()
        layout.addLayout(bh)

        self.cred_status = QLabel("")
        self.cred_status.setStyleSheet("color:#4c4; font-size:11px;")
        layout.addWidget(self.cred_status)

        layout.addStretch()
        tip = QLabel("🔒 Tip: Credentials are encrypted in transit and stored locally on your machine in config.ini.\nKeep the file in a safe location.")
        tip.setStyleSheet("color:#858585; font-size:10px;")
        tip.setWordWrap(True)
        tip.setMinimumWidth(0)
        layout.addWidget(tip)

        return w

    # =================================================================
    # TAB 3: SYSTEM & DEFAULTS
    # =================================================================
    def _create_system_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(12)

        # Paths Group
        g_paths = QGroupBox("Default File Locations")
        gp = QGridLayout(g_paths)
        gp.setHorizontalSpacing(10)
        gp.setVerticalSpacing(8)

        gp.addWidget(QLabel("Download Folder:"), 0, 0)
        self.dl_dir_ed = QLineEdit(r"D:\PITC_PDFs")
        gp.addWidget(self.dl_dir_ed, 0, 1)
        btn_dl = QPushButton("Browse...")
        btn_dl.clicked.connect(self._browse_dl_dir)
        gp.addWidget(btn_dl, 0, 2)

        gp.addWidget(QLabel("Chrome Profile:"), 1, 0)
        self.profile_dir_ed = QLineEdit(r"C:\ChromeProfiles\PITC")
        gp.addWidget(self.profile_dir_ed, 1, 1)
        btn_pr = QPushButton("Browse...")
        btn_pr.clicked.connect(self._browse_profile_dir)
        gp.addWidget(btn_pr, 1, 2)

        gp.setColumnStretch(1, 1)
        layout.addWidget(g_paths)

        # App Info Card
        lic = load_local_license()
        g_info = QGroupBox("Application and License Details")
        gi = QGridLayout(g_info)
        gi.setHorizontalSpacing(12)
        gi.setVerticalSpacing(6)

        gi.addWidget(QLabel("Suite Version:"), 0, 0)
        ver_lbl = QLabel(f"<b>{get_current_version()}</b> · Modern Edition")
        gi.addWidget(ver_lbl, 0, 1)

        gi.addWidget(QLabel("Plan Status:"), 1, 0)
        plan_lbl = QLabel(f"<span style='color:#4ec9b0; font-weight:bold;'>{format_plan_type(lic.get('plan_type'))}</span>")
        gi.addWidget(plan_lbl, 1, 1)

        gi.addWidget(QLabel("Expiration:"), 2, 0)
        exp_lbl = QLabel(format_expiry_date(lic.get('expires_at')))
        gi.addWidget(exp_lbl, 2, 1)

        gi.addWidget(QLabel("Hardware ID:"), 3, 0)
        hw_lbl = QLabel(f"<code>{get_device_id()}</code>")
        gi.addWidget(hw_lbl, 3, 1)

        gi.setColumnStretch(1, 1)
        layout.addWidget(g_info)

        bh = QHBoxLayout()
        btn_save_paths = QPushButton("Save Paths")
        btn_save_paths.setProperty("role", "primary")
        btn_save_paths.clicked.connect(self.save_paths)
        bh.addWidget(btn_save_paths)
        bh.addStretch()
        layout.addLayout(bh)

        self.system_status = QLabel("")
        self.system_status.setStyleSheet("color:#4c4; font-size:11px;")
        layout.addWidget(self.system_status)

        layout.addStretch()
        return w

    # =================================================================
    # ACTIONS & EVENT HANDLERS
    # =================================================================
    def _update_accent_buttons_style(self, theme_key: str = None):
        if theme_key is None:
            if hasattr(self, "theme_cb"):
                theme_key = self.theme_cb.currentData() or "vesper"
            else:
                from pitcsuite.ui_helpers import get_current_theme_key
                theme_key = get_current_theme_key()

        from pitcsuite.ui_helpers import THEMES, ACCENTS
        palette_data = THEMES.get(theme_key, THEMES["vesper"])
        is_light = (theme_key == "light")

        bg_unselected = palette_data.get("card_bg", "#ffffff") if is_light else palette_data.get("input_bg", "#252526")
        border_unselected = palette_data.get("border", "#d1d5db") if is_light else palette_data.get("border", "#3c3c3c")
        bg_hover = palette_data.get("item_hover", "#f1f5f9") if is_light else palette_data.get("item_hover", "#2e2e2e")

        light_text_map = {
            "amber": "#b45309",
            "teal": "#0e7490",
            "green": "#047857",
            "blue": "#006cc1",
            "purple": "#6d28d9",
            "red": "#be123c",
        }

        for key, btn in getattr(self, "accent_btns", {}).items():
            info = ACCENTS.get(key, ACCENTS["blue"])
            hex_c = info["hex"]
            text_c = light_text_map.get(key, hex_c) if is_light else hex_c

            btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: {bg_unselected};
                    color: {text_c};
                    border: 1px solid {border_unselected};
                    border-radius: 4px;
                    padding: 5px 8px;
                    font-weight: 600;
                    font-size: 8pt;
                }}
                QPushButton:hover {{
                    border-color: {hex_c};
                    background-color: {bg_hover};
                }}
                QPushButton:checked {{
                    background-color: {hex_c};
                    color: #ffffff;
                    border-color: {hex_c};
                }}
            """)

    def _select_accent(self, key: str):
        self._current_accent = key
        for k, btn in self.accent_btns.items():
            btn.setChecked(k == key)
        self._apply_current_ui(preview_only=True)

    def _on_live_preview_changed(self):
        if getattr(self, "_is_loading", False):
            return
        self._apply_current_ui(preview_only=True)

    def _apply_current_ui(self, preview_only=False):
        theme_key = self.theme_cb.currentData() or "vesper"
        accent_key = getattr(self, "_current_accent", "blue")
        font_family_key = self.font_cb.currentData() or "inter"
        font_size_key = self.scale_cb.currentData() or "standard"
        density_key = self.density_cb.currentData() or "compact"

        self._update_accent_buttons_style(theme_key)

        app = QApplication.instance()
        if app:
            apply_theme(app, theme_key, accent_key, font_size_key, density_key, font_family_key)

        # Notify parent MainWindow if active
        parent_win = self.window()
        if parent_win and hasattr(parent_win, "_on_theme_updated"):
            parent_win._on_theme_updated(theme_key, accent_key, font_size_key, density_key, font_family_key)

    def save_personalization(self):
        settings = QSettings("TNDevelopers", "PITCSuite")
        theme_key = self.theme_cb.currentData() or "vesper"
        accent_key = getattr(self, "_current_accent", "blue")
        font_family_key = self.font_cb.currentData() or "inter"
        scale_key = self.scale_cb.currentData() or "standard"
        density_key = self.density_cb.currentData() or "compact"
        sidebar_key = self.sidebar_style_cb.currentData() or "expanded"
        launch_key = self.launch_mode_cb.currentData() or "saved"

        settings.setValue("theme", theme_key)
        settings.setValue("accent_color", accent_key)
        settings.setValue("font_family", font_family_key)
        settings.setValue("ui_scale", scale_key)
        settings.setValue("density", density_key)
        settings.setValue("sidebar_style", sidebar_key)
        settings.setValue("launch_mode", launch_key)
        settings.setValue("smooth_anim", self.chk_smooth_anim.isChecked())
        settings.setValue("auto_update_check", self.chk_auto_update.isChecked())

        self._apply_current_ui()

        # Update sidebar only if target state differs from current
        parent_win = self.window()
        if parent_win and hasattr(parent_win, "nav_collapsed") and hasattr(parent_win, "set_sidebar_collapsed"):
            target_collapsed = (sidebar_key == "compact")
            if parent_win.nav_collapsed != target_collapsed:
                parent_win.set_sidebar_collapsed(target_collapsed)

        self.personalization_status.setText("✔  Personalization settings saved and applied!")
        self.personalization_status.setStyleSheet("color:#4ec9b0; font-size:11px;")

    def reset_personalization(self):
        idx_t = self.theme_cb.findData("vesper")
        if idx_t >= 0: self.theme_cb.setCurrentIndex(idx_t)

        self._select_accent("blue")

        idx_f = self.font_cb.findData("inter")
        if idx_f >= 0: self.font_cb.setCurrentIndex(idx_f)

        idx_s = self.scale_cb.findData("standard")
        if idx_s >= 0: self.scale_cb.setCurrentIndex(idx_s)

        idx_d = self.density_cb.findData("compact")
        if idx_d >= 0: self.density_cb.setCurrentIndex(idx_d)

        idx_sb = self.sidebar_style_cb.findData("expanded")
        if idx_sb >= 0: self.sidebar_style_cb.setCurrentIndex(idx_sb)

        idx_lm = self.launch_mode_cb.findData("saved")
        if idx_lm >= 0: self.launch_mode_cb.setCurrentIndex(idx_lm)

        self.chk_smooth_anim.setChecked(True)
        self.chk_auto_update.setChecked(True)

        self.save_personalization()
        self.personalization_status.setText("✔  Reset to default Vesper theme and styling.")

    def showEvent(self, event):
        super().showEvent(event)
        parent_win = self.window()
        if parent_win and hasattr(parent_win, "nav_collapsed"):
            curr_style = "compact" if parent_win.nav_collapsed else "expanded"
            idx = self.sidebar_style_cb.findData(curr_style)
            if idx >= 0 and self.sidebar_style_cb.currentIndex() != idx:
                self.sidebar_style_cb.blockSignals(True)
                self.sidebar_style_cb.setCurrentIndex(idx)
                self.sidebar_style_cb.blockSignals(False)
        self._update_accent_buttons_style()

    def load_all(self):
        self._is_loading = True
        try:
            settings = QSettings("TNDevelopers", "PITCSuite")

            # 1. Personalization (Defaulting to Vesper, Electric Blue, Inter, Standard Scale, Compact Density, Saved geometry)
            theme = str(settings.value("theme", "vesper")).lower()
            idx_t = self.theme_cb.findData(theme)
            if idx_t >= 0:
                self.theme_cb.setCurrentIndex(idx_t)

            accent = str(settings.value("accent_color", "blue")).lower()
            self._current_accent = accent
            if accent in self.accent_btns:
                self.accent_btns[accent].setChecked(True)

            font_family = str(settings.value("font_family", "inter")).lower()
            idx_f = self.font_cb.findData(font_family)
            if idx_f >= 0:
                self.font_cb.setCurrentIndex(idx_f)

            scale = str(settings.value("ui_scale", "standard")).lower()
            idx_s = self.scale_cb.findData(scale)
            if idx_s >= 0:
                self.scale_cb.setCurrentIndex(idx_s)

            density = str(settings.value("density", "compact")).lower()
            idx_d = self.density_cb.findData(density)
            if idx_d >= 0:
                self.density_cb.setCurrentIndex(idx_d)

            parent_win = self.window()
            if parent_win and hasattr(parent_win, "nav_collapsed"):
                sidebar = "compact" if parent_win.nav_collapsed else "expanded"
            else:
                sidebar = str(settings.value("sidebar_style", "expanded")).lower()
            idx_sb = self.sidebar_style_cb.findData(sidebar)
            if idx_sb >= 0:
                self.sidebar_style_cb.setCurrentIndex(idx_sb)

            launch = str(settings.value("launch_mode", "saved")).lower()
            idx_l = self.launch_mode_cb.findData(launch)
            if idx_l >= 0:
                self.launch_mode_cb.setCurrentIndex(idx_l)

            self.chk_smooth_anim.setChecked(settings.value("smooth_anim", True, type=bool))
            self.chk_auto_update.setChecked(settings.value("auto_update_check", True, type=bool))

            # 2. Credentials
            saved_disco = str(settings.value("pitc_disco", "PESCO")).strip() or "PESCO"
            path_cfg = _config_path()
            cfg = configparser.ConfigParser()
            if os.path.exists(path_cfg):
                cfg.read(path_cfg)
                self.user_ed.setText(cfg.get("credentials", "username", fallback=""))
                self.pass_ed.setText(cfg.get("credentials", "password", fallback=""))
                self.otp_ed.setText(cfg.get("credentials", "otp_secret", fallback=""))
                saved_disco = cfg.get("credentials", "disco", fallback=cfg.get("pitc_downloader", "disco", fallback=saved_disco))
                self.cred_status.setText("")
            else:
                self.cred_status.setText("⚠  config.ini not found — will be created on Save.")
                self.cred_status.setStyleSheet("color:#ca4; font-size:11px;")

            idx = self.disco_cb.findText(saved_disco)
            if idx >= 0:
                self.disco_cb.setCurrentIndex(idx)

            # 3. System Paths
            saved_dl = str(settings.value("default_download_dir", r"D:\PITC_PDFs")).strip() or r"D:\PITC_PDFs"
            saved_profile = str(settings.value("default_chrome_profile", r"C:\ChromeProfiles\PITC")).strip() or r"C:\ChromeProfiles\PITC"
            self.dl_dir_ed.setText(saved_dl)
            self.profile_dir_ed.setText(saved_profile)
        finally:
            self._is_loading = False
            self._update_accent_buttons_style()

    def save_credentials(self):
        username = self.user_ed.text().strip()
        password = self.pass_ed.text().strip()
        otp_secret = re.sub(r"[^A-Za-z2-7]", "", self.otp_ed.text()).upper()
        disco = self.disco_cb.currentText().strip() or "PESCO"

        if not username or not password:
            QMessageBox.warning(self, "Missing", "Username and password cannot be empty.")
            return

        settings = QSettings("TNDevelopers", "PITCSuite")
        settings.setValue("pitc_disco", disco)

        path = _config_path()
        cfg = configparser.ConfigParser()
        if os.path.exists(path):
            cfg.read(path)
        if not cfg.has_section("credentials"):
            cfg.add_section("credentials")
        cfg.set("credentials", "username", username)
        cfg.set("credentials", "password", password)
        cfg.set("credentials", "otp_secret", otp_secret)
        cfg.set("credentials", "disco", disco)
        with open(path, "w", encoding="utf-8") as f:
            cfg.write(f)

        self.cred_status.setText("✔  Credentials saved successfully.")
        self.cred_status.setStyleSheet("color:#4ec9b0; font-size:11px;")
        self.path_lbl.setText(f"Saved to:  {path}")

    def save_paths(self):
        dl = self.dl_dir_ed.text().strip()
        pr = self.profile_dir_ed.text().strip()
        settings = QSettings("TNDevelopers", "PITCSuite")
        settings.setValue("default_download_dir", dl)
        settings.setValue("default_chrome_profile", pr)
        self.system_status.setText("✔  Default paths saved successfully.")
        self.system_status.setStyleSheet("color:#4ec9b0; font-size:11px;")

    def _browse_dl_dir(self):
        p = QFileDialog.getExistingDirectory(self, "Select Default Download Folder", self.dl_dir_ed.text())
        if p: self.dl_dir_ed.setText(p)

    def _browse_profile_dir(self):
        p = QFileDialog.getExistingDirectory(self, "Select Chrome Profile Folder", self.profile_dir_ed.text())
        if p: self.profile_dir_ed.setText(p)

    # Backwards compatibility methods
    def load(self):
        self.load_all()

    def save(self):
        self.save_credentials()

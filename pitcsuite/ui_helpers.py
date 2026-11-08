import os
from PySide6.QtGui import QPalette, QColor, QFont
from PySide6.QtWidgets import QLabel, QFrame, QMessageBox, QPushButton, QApplication
from PySide6.QtCore import Qt, QSettings, QObject, QEvent

ASSETS_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "assets")).replace("\\", "/")

# =====================================================================
# THEME PALETTES & CONFIGURATION
# =====================================================================

THEMES = {
    "vesper": {
        "name": "Vesper (Default)",
        "window_bg": "#101010",
        "card_bg": "#161616",
        "card_border": "#262626",
        "input_bg": "#1b1b1b",
        "border": "#2e2e2e",
        "border_hover": "#424242",
        "nav_bg": "#0a0a0a",
        "nav_border": "#1c1c1c",
        "text_main": "#ffffff",
        "text_sec": "#cccccc",
        "text_muted": "#707070",
        "divider": "#1e1e1e",
        "item_hover": "#232323",
        "item_selected": "#2b2b2b",
        "tab_inactive": "#0a0a0a",
        "log_bg": "#0a0a0a",
        "badge_bg": "#181818",
    },
    "slate": {
        "name": "Slate (Cobalt Dark)",
        "window_bg": "#0d1117",
        "card_bg": "#161b22",
        "card_border": "#30363d",
        "input_bg": "#21262d",
        "border": "#30363d",
        "border_hover": "#484f58",
        "nav_bg": "#090d13",
        "nav_border": "#21262d",
        "text_main": "#f0f6fc",
        "text_sec": "#c9d1d9",
        "text_muted": "#8b949e",
        "divider": "#21262d",
        "item_hover": "#232a35",
        "item_selected": "#1f4068",
        "tab_inactive": "#090d13",
        "log_bg": "#090d13",
        "badge_bg": "#161b22",
    },
    "dark": {
        "name": "Dark",
        "window_bg": "#1e1e1e",
        "card_bg": "#252526",
        "card_border": "#333333",
        "input_bg": "#313131",
        "border": "#3c3c3c",
        "border_hover": "#4a4a4a",
        "nav_bg": "#181818",
        "nav_border": "#2b2b2b",
        "text_main": "#ffffff",
        "text_sec": "#cccccc",
        "text_muted": "#858585",
        "divider": "#2b2b2b",
        "item_hover": "#2a2d2e",
        "item_selected": "#04395e",
        "tab_inactive": "#181818",
        "log_bg": "#181818",
        "badge_bg": "#252526",
    },
    "midnight": {
        "name": "Midnight (OLED Dark)",
        "window_bg": "#09090b",
        "card_bg": "#121215",
        "card_border": "#222226",
        "input_bg": "#18181c",
        "border": "#27272a",
        "border_hover": "#3f3f46",
        "nav_bg": "#040405",
        "nav_border": "#1a1a1f",
        "text_main": "#ffffff",
        "text_sec": "#d4d4d8",
        "text_muted": "#71717a",
        "divider": "#1f1f23",
        "item_hover": "#202026",
        "item_selected": "#143354",
        "tab_inactive": "#040405",
        "log_bg": "#0d0d10",
        "badge_bg": "#121215",
    },
    "light": {
        "name": "Clean Light",
        "window_bg": "#f3f3f3",
        "card_bg": "#ffffff",
        "card_border": "#e5e7eb",
        "input_bg": "#ffffff",
        "border": "#d1d5db",
        "border_hover": "#9ca3af",
        "nav_bg": "#eaeaea",
        "nav_border": "#d1d5db",
        "text_main": "#111827",
        "text_sec": "#374151",
        "text_muted": "#6b7280",
        "divider": "#e5e7eb",
        "item_hover": "#e2e8f0",
        "item_selected": "#c7d2fe",
        "tab_inactive": "#e5e7eb",
        "log_bg": "#f8fafc",
        "badge_bg": "#ffffff",
    },
}

ACCENTS = {
    "blue": {
        "name": "Electric Blue",
        "hex": "#0078d4",
        "hover": "#006cc1",
        "pressed": "#005a9e",
    },
    "purple": {
        "name": "Royal Purple",
        "hex": "#7c3aed",
        "hover": "#6d28d9",
        "pressed": "#5b21b6",
    },
    "green": {
        "name": "Emerald Green",
        "hex": "#10b981",
        "hover": "#059669",
        "pressed": "#047857",
    },
    "amber": {
        "name": "Amber Gold",
        "hex": "#f59e0b",
        "hover": "#d97706",
        "pressed": "#b45309",
    },
    "red": {
        "name": "Crimson Red",
        "hex": "#e11d48",
        "hover": "#be123c",
        "pressed": "#9f1239",
    },
    "teal": {
        "name": "Cyan Teal",
        "hex": "#06b6d4",
        "hover": "#0891b2",
        "pressed": "#0e7490",
    },
}

FONTS = {
    "inter": {
        "name": "Inter",
        "family": "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
        "qt_family": "Inter",
    },
    "system": {
        "name": "System UI",
        "family": "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
        "qt_family": "Segoe UI",
    },
    "google_sans": {
        "name": "Google Sans",
        "family": "'Google Sans', 'Product Sans', 'Segoe UI', sans-serif",
        "qt_family": "Google Sans",
    },
}

FONT_SIZES = {
    "compact": 8.0,
    "standard": 8.5,
    "large": 9.5,
    "xlarge": 10.5,
}


def get_current_theme_key() -> str:
    settings = QSettings("TNDevelopers", "PITCSuite")
    return str(settings.value("theme", "vesper")).lower()


def get_current_accent_key() -> str:
    settings = QSettings("TNDevelopers", "PITCSuite")
    return str(settings.value("accent_color", "red")).lower()


def get_current_font_family_key() -> str:
    settings = QSettings("TNDevelopers", "PITCSuite")
    return str(settings.value("font_family", "inter")).lower()


def get_current_font_size_key() -> str:
    settings = QSettings("TNDevelopers", "PITCSuite")
    return str(settings.value("ui_scale", "compact")).lower()


def get_current_density_key() -> str:
    settings = QSettings("TNDevelopers", "PITCSuite")
    return str(settings.value("density", "compact")).lower()


def generate_theme_stylesheet(theme_key="vesper", accent_key="red", font_size_key="compact", density="compact", font_family_key="inter") -> str:
    palette_data = THEMES.get(theme_key, THEMES["vesper"])
    accent = ACCENTS.get(accent_key, ACCENTS["red"])
    font_pt = FONT_SIZES.get(font_size_key, 8.0)
    font_info = FONTS.get(font_family_key, FONTS["inter"])
    font_family_css = font_info["family"]
    
    is_compact = density == "compact"
    input_pad = "4px 7px" if is_compact else "6px 9px"
    cb_pad = "4px 24px 4px 7px" if is_compact else "6px 26px 6px 9px"
    input_min_h = "18px" if is_compact else "20px"
    btn_pad = "4px 10px" if is_compact else "6px 14px"
    group_pad = "8px 7px 7px 7px" if is_compact else "11px 9px 9px 9px"
    group_margin = "10px" if is_compact else "13px"

    is_light = theme_key == "light"
    chevron_down_asset = f"{ASSETS_DIR}/chevron_down_dark.svg" if is_light else f"{ASSETS_DIR}/chevron_down_light.svg"
    chevron_up_asset = f"{ASSETS_DIR}/chevron_up_dark.svg" if is_light else f"{ASSETS_DIR}/chevron_up_light.svg"
    checkmark_asset = f"{ASSETS_DIR}/checkmark.svg"

    css = f"""
/* Modern Theme: {palette_data['name']} / {accent['name']} */
QWidget {{
    color: {palette_data['text_sec']};
    font-family: {font_family_css};
    font-size: {font_pt}pt;
}}

QMainWindow, QDialog {{
    background-color: {palette_data['window_bg']};
}}

QWidget#contentArea {{
    background-color: {palette_data['window_bg']};
}}

QWidget#navSidebar {{
    background-color: {palette_data['nav_bg']};
    border-right: 1px solid {palette_data['nav_border']};
}}

/* Scroll Areas */
QScrollArea {{
    background: transparent;
    border: none;
}}

QScrollArea > QWidget > QWidget {{
    background: transparent;
}}

/* Dividers */
QFrame[frameShape="4"], QFrame[frameShape="5"] {{
    border: none;
    background-color: {palette_data['divider']};
    height: 1px;
    max-height: 1px;
}}

/* Containers / GroupBoxes */
QGroupBox {{
    background-color: {palette_data['card_bg']};
    border: 1px solid {palette_data['card_border']};
    border-radius: 6px;
    margin-top: {group_margin};
    padding: {group_pad};
    font-weight: 600;
}}

QGroupBox::title {{
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 10px;
    padding: 2px 8px;
    color: {palette_data['text_main']};
    background-color: {palette_data['card_bg']};
    border: 1px solid {palette_data['border']};
    border-radius: 4px;
    font-size: {font_pt - 0.5}pt;
    font-weight: 600;
}}

/* Inputs, Spinners, Dates */
QLineEdit, QSpinBox, QDoubleSpinBox, QDateEdit {{
    background-color: {palette_data['input_bg']};
    color: {palette_data['text_main']};
    border: 1px solid {palette_data['border']};
    border-radius: 4px;
    padding: {input_pad};
    font-size: {font_pt}pt;
    min-height: {input_min_h};
    selection-background-color: {accent['hex']};
    selection-color: #ffffff;
}}

QLineEdit:hover, QSpinBox:hover, QDoubleSpinBox:hover, QDateEdit:hover {{
    border-color: {palette_data['border_hover']};
}}

QLineEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus, QDateEdit:focus {{
    border: 1px solid {accent['hex']};
    background-color: {palette_data['input_bg']};
    color: {palette_data['text_main']};
}}

QLineEdit:read-only {{
    background-color: {palette_data['divider']};
    color: {palette_data['text_muted']};
    border-color: {palette_data['border']};
}}

QLineEdit:disabled, QSpinBox:disabled, QDoubleSpinBox:disabled, QDateEdit:disabled {{
    background-color: {palette_data['card_bg']};
    color: {palette_data['text_muted']};
    border-color: {palette_data['card_border']};
}}

/* SpinBox controls */
QSpinBox::up-button, QDoubleSpinBox::up-button, QDateEdit::up-button {{
    subcontrol-origin: border;
    subcontrol-position: top right;
    width: 16px;
    border-left: 1px solid {palette_data['border']};
    border-bottom: 1px solid {palette_data['border']};
    border-top-right-radius: 3px;
    background: {palette_data['card_bg']};
}}

QSpinBox::up-button:hover, QDoubleSpinBox::up-button:hover, QDateEdit::up-button:hover {{
    background: {accent['hex']};
}}

QSpinBox::up-arrow, QDoubleSpinBox::up-arrow, QDateEdit::up-arrow {{
    image: url("{chevron_up_asset}");
    width: 8px;
    height: 8px;
}}

QSpinBox::down-button, QDoubleSpinBox::down-button, QDateEdit::down-button {{
    subcontrol-origin: border;
    subcontrol-position: bottom right;
    width: 16px;
    border-left: 1px solid {palette_data['border']};
    border-bottom-right-radius: 3px;
    background: {palette_data['card_bg']};
}}

QSpinBox::down-button:hover, QDoubleSpinBox::down-button:hover, QDateEdit::down-button:hover {{
    background: {accent['hex']};
}}

QSpinBox::down-arrow, QDoubleSpinBox::down-arrow, QDateEdit::down-arrow {{
    image: url("{chevron_down_asset}");
    width: 8px;
    height: 8px;
}}

/* Combo Boxes */
QComboBox {{
    background-color: {palette_data['input_bg']};
    color: {palette_data['text_main']};
    border: 1px solid {palette_data['border']};
    border-radius: 4px;
    padding: {cb_pad};
    font-size: {font_pt}pt;
    min-height: {input_min_h};
}}

QComboBox:hover {{
    border-color: {palette_data['border_hover']};
}}

QComboBox:focus {{
    border: 1px solid {accent['hex']};
    background-color: {palette_data['input_bg']};
    color: {palette_data['text_main']};
}}

QComboBox::drop-down {{
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 20px;
    border-left: 1px solid {palette_data['border']};
    border-top-right-radius: 3px;
    border-bottom-right-radius: 3px;
    background: {palette_data['card_bg']};
}}

QComboBox::drop-down:hover {{
    background: {palette_data['border']};
}}

QComboBox::down-arrow {{
    image: url("{chevron_down_asset}");
    width: 12px;
    height: 12px;
}}

QComboBox QAbstractItemView {{
    background-color: {palette_data['card_bg']};
    color: {palette_data['text_main']};
    border: 1px solid {palette_data['border']};
    selection-background-color: {palette_data['item_selected']};
    selection-color: #ffffff;
    padding: 3px;
    border-radius: 4px;
    outline: none;
}}

/* Checkboxes and Radio Buttons */
QCheckBox, QRadioButton {{
    color: {palette_data['text_sec']};
    spacing: 7px;
    font-size: {font_pt}pt;
}}

QCheckBox:hover, QRadioButton:hover {{
    color: {palette_data['text_main']};
}}

QCheckBox::indicator, QRadioButton::indicator {{
    width: 15px;
    height: 15px;
    border: 1px solid {palette_data['border']};
    border-radius: 3px;
    background: {palette_data['input_bg']};
}}

QRadioButton::indicator {{
    border-radius: 7px;
}}

QCheckBox::indicator:hover, QRadioButton::indicator:hover {{
    border-color: {accent['hex']};
    background: {palette_data['card_bg']};
}}

QCheckBox::indicator:checked {{
    background-color: {accent['hex']};
    border-color: {accent['hex']};
    image: url("{checkmark_asset}");
}}

QRadioButton::indicator:checked {{
    background-color: {accent['hex']};
    border: 3px solid {palette_data['input_bg']};
}}

/* Progress Bars */
QProgressBar {{
    background-color: {palette_data['card_bg']};
    border: 1px solid {palette_data['card_border']};
    border-radius: 4px;
    text-align: center;
    color: #ffffff;
    font-size: {font_pt - 1}pt;
    font-weight: 600;
    min-height: 12px;
    max-height: 12px;
}}

QProgressBar::chunk {{
    background-color: {accent['hex']};
    border-radius: 3px;
}}

/* Terminal & Log Text Areas */
QTextEdit, QPlainTextEdit {{
    background-color: {palette_data['log_bg']};
    color: {palette_data['text_sec']};
    border: 1px solid {palette_data['divider']};
    border-radius: 4px;
    padding: 7px;
    font-family: 'Cascadia Code', 'Consolas', 'Google Sans Mono', monospace;
    font-size: {font_pt - 0.5}pt;
    selection-background-color: {accent['hex']};
    selection-color: #ffffff;
}}

/* Lists and Tables */
QListWidget, QTableWidget, QTableView, QTreeView {{
    background-color: {palette_data['log_bg']};
    color: {palette_data['text_sec']};
    border: 1px solid {palette_data['divider']};
    border-radius: 4px;
    gridline-color: {palette_data['card_bg']};
    selection-background-color: {palette_data['item_selected']};
    selection-color: #ffffff;
    outline: none;
}}

QListWidget::item, QTableWidget::item {{
    padding: 4px 6px;
    border-radius: 3px;
}}

QListWidget::item:hover, QTableWidget::item:hover {{
    background-color: {palette_data['item_hover']};
}}

QListWidget::item:selected, QTableWidget::item:selected {{
    background-color: {palette_data['item_selected']};
    color: #ffffff;
}}

QHeaderView::section {{
    background-color: {palette_data['card_bg']};
    color: {palette_data['text_muted']};
    border: none;
    border-bottom: 1px solid {palette_data['divider']};
    border-right: 1px solid {palette_data['divider']};
    padding: 5px 6px;
    font-weight: 600;
    font-size: {font_pt - 0.5}pt;
}}

QTableCornerButton::section {{
    background-color: {palette_data['card_bg']};
    border: none;
}}

/* Antigravity / VS Code Scrollbars */
QScrollBar:vertical {{
    background: transparent;
    width: 8px;
    margin: 0;
}}

QScrollBar::handle:vertical {{
    background: rgba(121, 121, 121, 0.35);
    min-height: 24px;
    border-radius: 3px;
}}

QScrollBar::handle:vertical:hover {{
    background: rgba(100, 100, 100, 0.65);
}}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical,
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{
    height: 0px;
    background: none;
}}

QScrollBar:horizontal {{
    background: transparent;
    height: 8px;
    margin: 0;
}}

QScrollBar::handle:horizontal {{
    background: rgba(121, 121, 121, 0.35);
    min-width: 24px;
    border-radius: 3px;
}}

QScrollBar::handle:horizontal:hover {{
    background: rgba(100, 100, 100, 0.65);
}}

QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal,
QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {{
    width: 0px;
    background: none;
}}

/* Modern Antigravity Tabs */
QTabWidget::pane {{
    border: 1px solid {palette_data['card_border']};
    background-color: {palette_data['card_bg']};
    border-radius: 6px;
}}

QTabBar::tab {{
    background: {palette_data['tab_inactive']};
    color: {palette_data['text_muted']};
    padding: 7px 16px;
    border-top-left-radius: 5px;
    border-top-right-radius: 5px;
    margin-right: 2px;
    border: 1px solid {palette_data['divider']};
    border-bottom: none;
    font-weight: 500;
    font-size: {font_pt}pt;
}}

QTabBar::tab:selected {{
    background: {palette_data['card_bg']};
    color: {palette_data['text_main']};
    font-weight: 600;
    border-bottom: 2px solid {accent['hex']};
}}

QTabBar::tab:hover:!selected {{
    background: {palette_data['item_hover']};
    color: {palette_data['text_sec']};
}}

/* Tooltips */
QToolTip {{
    background-color: {palette_data['card_bg']};
    color: {palette_data['text_sec']};
    border: 1px solid {palette_data['border']};
    border-radius: 4px;
    padding: 4px 8px;
    font-size: {font_pt - 0.5}pt;
}}

/* Antigravity Standard Secondary PushButton */
QPushButton {{
    background: {palette_data['input_bg']};
    color: {palette_data['text_main']};
    border: 1px solid {palette_data['border']};
    border-radius: 4px;
    padding: {btn_pad};
    font-size: {font_pt}pt;
    font-weight: 500;
}}

QPushButton:hover {{
    background: {palette_data['card_border']};
    border-color: {palette_data['border_hover']};
    color: {palette_data['text_main']};
}}

QPushButton:pressed {{
    background: {palette_data['card_bg']};
}}

QPushButton:disabled {{
    background: {palette_data['card_bg']};
    color: {palette_data['text_muted']};
    border-color: {palette_data['card_border']};
}}

/* Primary role button */
QPushButton[role=primary] {{
    background: {accent['hex']};
    color: #ffffff;
    border: 1px solid {accent['hex']};
    border-radius: 4px;
    padding: {btn_pad};
    font-weight: 600;
    font-size: {font_pt}pt;
}}

QPushButton[role=primary]:hover {{
    background: {accent['hover']};
    border-color: {accent['hover']};
}}

QPushButton[role=primary]:pressed {{
    background: {accent['pressed']};
}}

/* Secondary role button */
QPushButton[role=secondary] {{
    background: {palette_data['input_bg']};
    color: {palette_data['text_main']};
    border: 1px solid {palette_data['border']};
    border-radius: 4px;
    padding: {btn_pad};
    font-weight: 500;
}}

QPushButton[role=secondary]:hover {{
    background: {palette_data['card_border']};
    border-color: {palette_data['border_hover']};
    color: {palette_data['text_main']};
}}
"""
    return css


def get_start_btn_style(accent_key=None) -> str:
    if accent_key is None:
        accent_key = get_current_accent_key()
    accent = ACCENTS.get(accent_key, ACCENTS["blue"])
    return f"""
QPushButton {{
    background-color: {accent['hex']};
    color: #ffffff;
    border: 1px solid {accent['hex']};
    border-radius: 4px;
    padding: 6px 18px;
    font-weight: 600;
    font-size: 8.5pt;
}}
QPushButton:hover {{
    background-color: {accent['hover']};
    border-color: {accent['hover']};
}}
QPushButton:pressed {{
    background-color: {accent['pressed']};
    border-color: {accent['pressed']};
}}
QPushButton:disabled {{
    background-color: #222222;
    color: #656565;
    border: 1px solid #2d2d2d;
}}
"""


def get_stop_btn_style() -> str:
    return """
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


def get_nav_style(accent_key=None, theme_key=None) -> str:
    if accent_key is None:
        accent_key = get_current_accent_key()
    if theme_key is None:
        theme_key = get_current_theme_key()
    accent = ACCENTS.get(accent_key, ACCENTS["blue"])
    palette_data = THEMES.get(theme_key, THEMES["dark"])
    
    return f"""
QPushButton {{
    background: transparent;
    color: {palette_data['text_muted']};
    border: none;
    border-left: 2px solid transparent;
    border-radius: 0px;
    padding: 6px 10px;
    text-align: left;
    font-size: 8.5pt;
    font-weight: 500;
    margin: 1px 0px;
}}
QPushButton:hover {{
    background: {palette_data['item_hover']};
    color: {palette_data['text_main']};
}}
QPushButton[active="true"] {{
    background: {palette_data['card_bg']};
    color: {palette_data['text_main']};
    border-left: 2px solid {accent['hex']};
    font-weight: 600;
}}
"""


# Static aliases for backwards compatibility
DARK_STYLE = generate_theme_stylesheet()
START_BTN = get_start_btn_style()
STOP_BTN = get_stop_btn_style()
NAV_STYLE = get_nav_style()


class EscMessageBoxFilter(QObject):
    """
    Global application event filter that ensures pressing the Escape key (Esc)
    automatically closes and dismisses any QMessageBox or modal dialog across all tools.
    """
    def eventFilter(self, watched, event):
        if event.type() == QEvent.KeyPress and event.key() == Qt.Key_Escape:
            widget = None
            if isinstance(watched, QMessageBox):
                widget = watched
            elif hasattr(watched, "window") and isinstance(watched.window(), QMessageBox):
                widget = watched.window()
            else:
                active = QApplication.activeModalWidget() or QApplication.activeWindow()
                if isinstance(active, QMessageBox):
                    widget = active

            if widget is not None:
                esc_btn = widget.escapeButton()
                if esc_btn:
                    esc_btn.click()
                    return True
                for sb in [QMessageBox.Cancel, QMessageBox.No, QMessageBox.Close, QMessageBox.Discard, QMessageBox.Ok]:
                    btn = widget.button(sb)
                    if btn:
                        btn.click()
                        return True
                if widget.buttons():
                    widget.buttons()[-1].click()
                    return True
                widget.reject()
                return True
        return super().eventFilter(watched, event)


def configure_app(app, theme_key=None, accent_key=None, font_size_key=None, density=None, font_family_key=None):
    if not hasattr(app, "_esc_filter_installed"):
        app._esc_filter = EscMessageBoxFilter(app)
        app.installEventFilter(app._esc_filter)
        app._esc_filter_installed = True
    if theme_key is None:
        theme_key = get_current_theme_key()
    if accent_key is None:
        accent_key = get_current_accent_key()
    if font_size_key is None:
        font_size_key = get_current_font_size_key()
    if density is None:
        density = get_current_density_key()
    if font_family_key is None:
        font_family_key = get_current_font_family_key()

    palette_data = THEMES.get(theme_key, THEMES["vesper"])
    accent = ACCENTS.get(accent_key, ACCENTS["red"])
    font_pt = FONT_SIZES.get(font_size_key, 8.0)
    font_info = FONTS.get(font_family_key, FONTS["inter"])

    app.setStyle("Fusion")

    font = app.font()
    font.setFamily(font_info["qt_family"])
    font.setPointSizeF(font_pt)
    app.setFont(font)

    palette = QPalette()
    colors = {
        "Window": palette_data["window_bg"],
        "WindowText": palette_data["text_sec"],
        "Base": palette_data["nav_bg"],
        "AlternateBase": palette_data["card_bg"],
        "ToolTipBase": palette_data["card_bg"],
        "ToolTipText": palette_data["text_sec"],
        "Text": palette_data["text_main"],
        "Button": palette_data["input_bg"],
        "ButtonText": palette_data["text_main"],
        "BrightText": "#ffffff",
        "Highlight": accent["hex"],
        "HighlightedText": "#ffffff",
        "PlaceholderText": palette_data["text_muted"],
        "Link": accent["hex"],
    }
    for role, color in colors.items():
        palette.setColor(getattr(QPalette.ColorRole, role), QColor(color))

    for role in (QPalette.ColorRole.Text, QPalette.ColorRole.WindowText, QPalette.ColorRole.ButtonText):
        palette.setColor(QPalette.ColorGroup.Disabled, role, QColor(palette_data["text_muted"]))

    app.setPalette(palette)


def apply_theme(app, theme_key=None, accent_key=None, font_size_key=None, density=None, font_family_key=None):
    """Dynamically applies the chosen theme and re-styles the entire application in real time."""
    if theme_key is None:
        theme_key = get_current_theme_key()
    if accent_key is None:
        accent_key = get_current_accent_key()
    if font_size_key is None:
        font_size_key = get_current_font_size_key()
    if density is None:
        density = get_current_density_key()
    if font_family_key is None:
        font_family_key = get_current_font_family_key()

    configure_app(app, theme_key, accent_key, font_size_key, density, font_family_key)
    stylesheet = generate_theme_stylesheet(theme_key, accent_key, font_size_key, density, font_family_key)
    app.setStyleSheet(stylesheet)

    # Update global module references so panels creating widgets later get updated styles
    global DARK_STYLE, START_BTN, STOP_BTN, NAV_STYLE
    DARK_STYLE = stylesheet
    START_BTN = get_start_btn_style(accent_key)
    STOP_BTN = get_stop_btn_style()
    NAV_STYLE = get_nav_style(accent_key, theme_key)


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
    box.addButton(QMessageBox.Ok)
    box.setEscapeButton(QMessageBox.Ok)

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

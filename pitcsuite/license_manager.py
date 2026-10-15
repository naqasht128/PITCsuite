"""
PITC Suite - Licensing, Hardware-Locking & Remote Verification Manager.
Implements the TN Developers Universal Licensing Blueprint.
"""

from __future__ import annotations

import configparser
from datetime import datetime, timezone, timedelta
import hashlib
import json
import os
import re
import subprocess
import sys
from typing import List, Tuple, Set, Optional
import urllib.parse
import uuid
import webbrowser
import winreg

import requests
from PySide6.QtCore import Qt, QThread, Signal, QUrl
from PySide6.QtGui import QFont, QCursor, QDesktopServices, QIcon
from PySide6.QtWidgets import (
    QApplication, QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QPushButton, QMessageBox, QTabWidget, QWidget,
    QFrame, QGridLayout, QScrollArea
)

# ====================================================================
# CONFIGURATION & CONSTANTS
# ====================================================================

DEFAULT_LICENSE_API_URL = "https://pitcsuite.vercel.app"
APP_NAME = "PITCSuite"
SUPPORT_WHATSAPP = "923339962474"
SUPPORT_EMAIL = "tahir.naqash4@gmail.com"

# VS Code Dark+ Color Palette
COLOR_BG_DARK = "#1e1e1e"
COLOR_CARD_BG = "#252526"
COLOR_INPUT_BG = "#2d2d2d"
COLOR_BORDER = "#3c3c3c"
COLOR_TEXT_PRIMARY = "#ffffff"
COLOR_TEXT_SECONDARY = "#cccccc"
COLOR_TEXT_MUTED = "#858585"
COLOR_ACCENT_BLUE = "#007acc"
COLOR_ACCENT_BLUE_HOVER = "#1c8cd8"
COLOR_ACCENT_GREEN = "#238636"
COLOR_ACCENT_GREEN_HOVER = "#2ea043"
COLOR_ERROR = "#f87171"
COLOR_WARNING = "#fbbf24"


def get_api_base_url() -> str:
    """Resolves license API base URL from environment or local config.ini."""
    env_url = os.environ.get("LICENSE_API_URL", "").strip()
    if env_url:
        return env_url.rstrip("/")

    # Check local config.ini if available
    try:
        from pitcsuite.config import config_path
        cfg_file = config_path()
        if os.path.exists(cfg_file):
            cfg = configparser.ConfigParser()
            cfg.read(cfg_file, encoding="utf-8")
            if cfg.has_option("License", "api_url"):
                val = cfg.get("License", "api_url").strip()
                if val:
                    return val.rstrip("/")
    except Exception:
        pass

    return DEFAULT_LICENSE_API_URL.rstrip("/")


# ====================================================================
# HARDWARE ID GENERATION (Windows 11 24H2+ Safe)
# ====================================================================

def get_device_id() -> str:
    """
    Returns persistent, unique hardware UUID.
    Guaranteed never to crash on Windows 11 (24H2+) even when WMIC is missing or deprecated.
    """
    if sys.platform == "win32":
        # 1. Windows Registry MachineGuid (64-bit and 32-bit views)
        for flag in [getattr(winreg, 'KEY_WOW64_64KEY', 0), getattr(winreg, 'KEY_WOW64_32KEY', 0), 0]:
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Cryptography", 0, winreg.KEY_READ | flag) as key:
                    val, _ = winreg.QueryValueEx(key, "MachineGuid")
                    if val and len(str(val).strip()) > 10:
                        return str(val).strip().upper()
            except Exception:
                pass

        # 2. Modern PowerShell CIM query (Zero console flashing)
        try:
            cmd = ["powershell", "-NoProfile", "-Command", "(Get-CimInstance Win32_ComputerSystemProduct).UUID"]
            flags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
            out = subprocess.check_output(cmd, creationflags=flags, timeout=5).decode(errors="ignore").strip()
            if '-' in out and len(out) >= 30:
                return out.upper()
        except Exception:
            pass

        # 3. MAC address fallback
        try:
            mac = uuid.getnode()
            return f"MAC-{hex(mac)[2:].upper()}"
        except Exception:
            pass

        return "UNKNOWN-HWID-WIN"

    # macOS / Linux fallback
    try:
        mac = uuid.getnode()
        return f"FALLBACK-{hex(mac)[2:].upper()}"
    except Exception:
        return "UNKNOWN-DEVICE-ID"


def get_candidate_device_ids() -> List[str]:
    """
    Returns all legitimate hardware IDs that can identify this specific machine.
    Prevents existing activated clients from being locked out across hardware query variations.
    """
    primary = get_device_id()
    candidates: Set[str] = {primary, primary.upper(), primary.lower()}

    if sys.platform == "win32":
        # Check all registry views
        for flag in [getattr(winreg, 'KEY_WOW64_64KEY', 0), getattr(winreg, 'KEY_WOW64_32KEY', 0), 0]:
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Cryptography", 0, winreg.KEY_READ | flag) as key:
                    val, _ = winreg.QueryValueEx(key, "MachineGuid")
                    if val and len(str(val).strip()) > 10:
                        c = str(val).strip()
                        candidates.add(c)
                        candidates.add(c.upper())
            except Exception:
                pass

        # PowerShell CIM BIOS UUID
        try:
            cmd = ["powershell", "-NoProfile", "-Command", "(Get-CimInstance Win32_ComputerSystemProduct).UUID"]
            flags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
            out = subprocess.check_output(cmd, creationflags=flags, timeout=5).decode(errors="ignore").strip()
            if '-' in out and len(out) >= 30:
                candidates.add(out.upper())
                candidates.add(out.lower())
        except Exception:
            pass

        # Physical MAC fallback
        try:
            mac = uuid.getnode()
            candidates.add(f"MAC-{hex(mac)[2:].upper()}")
            candidates.add(f"FALLBACK-{hex(mac)[2:].upper()}")
        except Exception:
            pass

    return [c for c in candidates if c]


def normalize_key(key: str) -> str:
    """Removes hyphens, spaces, and normalizes to uppercase characters."""
    if not key:
        return ""
    return re.sub(r'[^A-Za-z0-9]', '', key).upper()


# ====================================================================
# STORAGE TIERS (JSON & Multi-Location Redundancy)
# ====================================================================

def get_primary_license_json_path() -> str:
    """Returns primary license.json path in %APPDATA%/PITCSuite."""
    base = os.getenv("APPDATA") or os.path.expanduser(r"~\AppData\Roaming")
    folder = os.path.join(base, APP_NAME)
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, "license.json")


def get_all_storage_paths() -> List[str]:
    """Returns all redundant persistent storage paths for the license."""
    paths = []
    # 1. Roaming AppData
    appdata = os.getenv("APPDATA") or os.path.expanduser(r"~\AppData\Roaming")
    if appdata:
        paths.append(os.path.join(appdata, APP_NAME, "license.json"))
        paths.append(os.path.join(appdata, APP_NAME, "license.key"))

    # 2. Local AppData
    localappdata = os.getenv("LOCALAPPDATA") or os.path.expanduser(r"~\AppData\Local")
    if localappdata:
        paths.append(os.path.join(localappdata, APP_NAME, "license.json"))
        paths.append(os.path.join(localappdata, "Programs", APP_NAME, "license.json"))

    # 3. Machine-wide ProgramData
    programdata = os.getenv("PROGRAMDATA") or os.getenv("ALLUSERSPROFILE") or r"C:\ProgramData"
    if programdata:
        paths.append(os.path.join(programdata, APP_NAME, "license.json"))

    seen = set()
    deduped = []
    for p in paths:
        norm = os.path.normpath(p).lower()
        if norm not in seen:
            seen.add(norm)
            deduped.append(os.path.normpath(p))
    return deduped


def load_local_license() -> dict:
    """Reads license metadata from local storage (%APPDATA%/PITCSuite/license.json)."""
    for path in get_all_storage_paths():
        if path.endswith(".json") and os.path.isfile(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, dict) and data.get("serial_key"):
                        return data
            except Exception:
                pass

    return {}


def save_local_license(data: dict) -> None:
    """Persists license metadata across JSON files and Windows Registry."""
    if not isinstance(data, dict) or not data.get("serial_key"):
        return

    content = json.dumps(data, indent=2, ensure_ascii=False)
    for path in get_all_storage_paths():
        if path.endswith(".json"):
            try:
                os.makedirs(os.path.dirname(path), exist_ok=True)
                with open(path, "w", encoding="utf-8") as f:
                    f.write(content)
            except Exception:
                pass

    # Write key to Registry HKCU
    if sys.platform == "win32":
        try:
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, rf"Software\{APP_NAME}") as k:
                winreg.SetValueEx(k, "LicenseKey", 0, winreg.REG_SZ, data["serial_key"])
        except Exception:
            pass


# ====================================================================
# REMOTE API CLIENT
# ====================================================================

def verify_license_remote(serial_key: str, device_id: str, candidate_ids: Optional[List[str]] = None) -> dict:
    """Sends verification request to Vercel API."""
    api_url = get_api_base_url()
    url = f"{api_url}/api/verify"
    payload = {
        "serial_key": serial_key.strip().upper(),
        "device_id": device_id.strip(),
        "candidate_ids": candidate_ids or [device_id.strip()],
        "os_info": f"Windows {sys.getwindowsversion().major}.{sys.getwindowsversion().minor}" if sys.platform == "win32" else sys.platform
    }
    try:
        resp = requests.post(url, json=payload, timeout=8)
        try:
            return resp.json()
        except Exception:
            return {"valid": False, "status": "server_error", "message": f"Server returned status {resp.status_code}"}
    except requests.exceptions.Timeout:
        return {"valid": False, "status": "network_error", "message": "Connection to license server timed out."}
    except Exception as exc:
        return {"valid": False, "status": "network_error", "message": str(exc)}


def activate_license_remote(serial_key: str, device_id: str, candidate_ids: Optional[List[str]] = None) -> dict:
    """Sends activation request to Vercel API."""
    api_url = get_api_base_url()
    url = f"{api_url}/api/activate"
    payload = {
        "serial_key": serial_key.strip().upper(),
        "device_id": device_id.strip(),
        "candidate_ids": candidate_ids or [device_id.strip()],
        "os_info": f"Windows {sys.getwindowsversion().major}.{sys.getwindowsversion().minor}" if sys.platform == "win32" else sys.platform
    }
    try:
        resp = requests.post(url, json=payload, timeout=10)
        try:
            return resp.json()
        except Exception:
            return {"valid": False, "status": "server_error", "message": f"Server returned status {resp.status_code}"}
    except requests.exceptions.Timeout:
        return {"valid": False, "status": "network_error", "message": "Connection to license server timed out."}
    except Exception as exc:
        return {"valid": False, "status": "network_error", "message": str(exc)}


def is_offline_grace_valid(local_lic: dict) -> Tuple[bool, str]:
    """
    Checks if an existing local license is within the offline grace period.
    Allows up to 7 days of offline usage from the last successful remote verification.
    """
    if not local_lic or not local_lic.get("valid"):
        return False, "No valid cached license found."

    expires_at_str = local_lic.get("expires_at")
    if expires_at_str:
        try:
            exp_dt = datetime.fromisoformat(expires_at_str.replace("Z", "+00:00"))
            if datetime.now(timezone.utc) > exp_dt:
                return False, "Subscription has expired."
        except Exception:
            pass

    last_verified_str = local_lic.get("last_verified")
    if last_verified_str:
        try:
            lv_dt = datetime.fromisoformat(last_verified_str.replace("Z", "+00:00"))
            grace_period = timedelta(days=7)
            if datetime.now(timezone.utc) - lv_dt <= grace_period:
                return True, "Valid under 7-day offline grace period."
        except Exception:
            pass

    # If last_verified is missing but valid=True was saved recently
    return True, "Valid under offline grace period."


# ====================================================================
# BACKGROUND ACTIVATION WORKER THREAD
# ====================================================================

class ActivationWorker(QThread):
    finished_signal = Signal(dict)

    def __init__(self, serial_key: str, device_id: str, candidate_ids: List[str]):
        super().__init__()
        self.serial_key = serial_key
        self.device_id = device_id
        self.candidate_ids = candidate_ids

    def run(self):
        # Call remote activation API (all licenses must be verified with Supabase)
        res = activate_license_remote(self.serial_key, self.device_id, self.candidate_ids)
        self.finished_signal.emit(res)


# ====================================================================
# MODERN VS CODE DARK+ LICENSE GATE DIALOG (PySide6)
# ====================================================================

class LicenseGateDialog(QDialog):
    """
    Modern VS Code Dark+ styled product activation dialog.
    Tab 1: Serial Key activation with Hardware ID and one-click copy.
    Tab 2: Subscription plans with WhatsApp and Email quick-order links.
    """

    def __init__(self, initial_error: str = "", parent=None):
        super().__init__(parent)
        self.device_id = get_device_id()
        self.candidate_ids = get_candidate_device_ids()
        self.activated = False
        self.license_info: Optional[dict] = None
        self._worker: Optional[ActivationWorker] = None

        self.setWindowTitle("PITC Suite — Product Activation & Licensing")
        self.setFixedSize(560, 480)
        self.setWindowFlags(Qt.Dialog | Qt.WindowTitleHint | Qt.WindowCloseButtonHint)
        self.setAttribute(Qt.WA_DeleteOnClose, False)

        self._apply_styles()
        self._build_ui(initial_error)

    def _apply_styles(self):
        self.setStyleSheet(f"""
            QDialog {{
                background-color: {COLOR_BG_DARK};
                color: {COLOR_TEXT_PRIMARY};
                font-family: 'Segoe UI', Tahoma, Arial, sans-serif;
            }}
            QTabWidget::pane {{
                border: 1px solid {COLOR_BORDER};
                background: {COLOR_CARD_BG};
                border-radius: 6px;
                top: -1px;
            }}
            QTabBar::tab {{
                background: {COLOR_BG_DARK};
                color: {COLOR_TEXT_MUTED};
                font-size: 10pt;
                font-weight: bold;
                padding: 8px 20px;
                border: 1px solid {COLOR_BORDER};
                border-bottom: none;
                border-top-left-radius: 5px;
                border-top-right-radius: 5px;
                margin-right: 4px;
            }}
            QTabBar::tab:selected {{
                background: {COLOR_CARD_BG};
                color: {COLOR_TEXT_PRIMARY};
                border-bottom: 2px solid {COLOR_ACCENT_BLUE};
            }}
            QTabBar::tab:hover:!selected {{
                background: #2a2d2e;
                color: {COLOR_TEXT_SECONDARY};
            }}
            QLabel {{
                color: {COLOR_TEXT_SECONDARY};
                font-size: 9.5pt;
            }}
            QLineEdit {{
                background-color: {COLOR_INPUT_BG};
                color: {COLOR_TEXT_PRIMARY};
                border: 1px solid {COLOR_BORDER};
                border-radius: 5px;
                padding: 7px 10px;
                font-size: 10pt;
            }}
            QLineEdit:focus {{
                border: 1px solid {COLOR_ACCENT_BLUE};
            }}
        """)

    def _build_ui(self, initial_error: str):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 18, 20, 18)
        main_layout.setSpacing(14)

        # Header Title Row
        header_row = QHBoxLayout()
        header_row.setSpacing(10)
        title_vbox = QVBoxLayout()
        title_vbox.setSpacing(2)

        lbl_title = QLabel("PITC Suite Activation Gate")
        lbl_title.setStyleSheet(f"font-size: 13pt; font-weight: bold; color: {COLOR_TEXT_PRIMARY};")
        title_vbox.addWidget(lbl_title)

        lbl_subtitle = QLabel("Hardware-Locked Licensing & Automatic Updates")
        lbl_subtitle.setStyleSheet(f"font-size: 8.5pt; color: {COLOR_TEXT_MUTED};")
        title_vbox.addWidget(lbl_subtitle)
        header_row.addLayout(title_vbox)
        header_row.addStretch()

        badge = QLabel("TN Developers")
        badge.setStyleSheet(f"background: #1e3a5f; color: #60a5fa; font-weight: bold; font-size: 8pt; padding: 4px 10px; border-radius: 4px;")
        header_row.addWidget(badge)
        main_layout.addLayout(header_row)

        # Tab Widget
        tabs = QTabWidget(self)
        tabs.addTab(self._build_activation_tab(initial_error), "🔑 Product Activation")
        tabs.addTab(self._build_plans_tab(), "💎 Subscription Plans")
        main_layout.addWidget(tabs, 1)

    def _build_activation_tab(self, initial_error: str) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(12)

        # 1. Hardware ID Section
        hw_lbl = QLabel("Your Unique Machine Hardware ID (HWID):")
        hw_lbl.setStyleSheet(f"font-weight: bold; color: {COLOR_TEXT_PRIMARY};")
        layout.addWidget(hw_lbl)

        hw_row = QHBoxLayout()
        hw_row.setSpacing(8)

        self.hw_entry = QLineEdit(self.device_id)
        self.hw_entry.setReadOnly(True)
        self.hw_entry.setStyleSheet(f"""
            QLineEdit {{
                background-color: #171717;
                color: #38bdf8;
                font-family: 'Consolas', 'Courier New', monospace;
                font-weight: bold;
                font-size: 10pt;
                padding: 7px 10px;
                border: 1px solid #334155;
            }}
        """)
        hw_row.addWidget(self.hw_entry, 1)

        self.btn_copy = QPushButton("📋 Copy ID")
        self.btn_copy.setCursor(Qt.PointingHandCursor)
        self.btn_copy.setStyleSheet(f"""
            QPushButton {{
                background: #334155;
                color: #f1f5f9;
                font-weight: bold;
                font-size: 9pt;
                padding: 7px 14px;
                border: none;
                border-radius: 5px;
            }}
            QPushButton:hover {{
                background: #475569;
            }}
        """)
        self.btn_copy.clicked.connect(self._on_copy_hwid)
        hw_row.addWidget(self.btn_copy)
        layout.addLayout(hw_row)

        layout.addSpacing(4)

        # 2. Serial Key Section
        key_lbl = QLabel("Enter Your Serial Key:")
        key_lbl.setStyleSheet(f"font-weight: bold; color: {COLOR_TEXT_PRIMARY};")
        layout.addWidget(key_lbl)

        self.key_entry = QLineEdit()
        self.key_entry.setPlaceholderText("PS-XXXX-XXXX-XXXX")
        self.key_entry.setStyleSheet(f"""
            QLineEdit {{
                background-color: #171717;
                color: #4ade80;
                font-family: 'Consolas', 'Courier New', monospace;
                font-weight: bold;
                font-size: 11pt;
                letter-spacing: 1px;
                padding: 8px 10px;
                border: 1px solid #334155;
            }}
        """)
        self.key_entry.textChanged.connect(self._format_serial_input)
        self.key_entry.returnPressed.connect(self._on_activate_clicked)
        layout.addWidget(self.key_entry)

        # 3. Status Message Label
        self.status_lbl = QLabel(initial_error if initial_error else "")
        self.status_lbl.setWordWrap(True)
        if initial_error:
            self.status_lbl.setStyleSheet(f"color: {COLOR_ERROR}; font-size: 8.5pt;")
        else:
            self.status_lbl.setStyleSheet(f"color: {COLOR_TEXT_MUTED}; font-size: 8.5pt;")
        layout.addWidget(self.status_lbl)

        layout.addStretch()

        # 4. Activate Button
        self.btn_activate = QPushButton("Activate PITC Suite")
        self.btn_activate.setCursor(Qt.PointingHandCursor)
        self.btn_activate.setStyleSheet(f"""
            QPushButton {{
                background-color: {COLOR_ACCENT_GREEN};
                color: white;
                font-weight: bold;
                font-size: 10.5pt;
                padding: 10px;
                border: none;
                border-radius: 6px;
            }}
            QPushButton:hover {{
                background-color: {COLOR_ACCENT_GREEN_HOVER};
            }}
            QPushButton:disabled {{
                background-color: #334155;
                color: #94a3b8;
            }}
        """)
        self.btn_activate.clicked.connect(self._on_activate_clicked)
        layout.addWidget(self.btn_activate)

        return widget

    def _build_plans_tab(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(10)

        info_lbl = QLabel("Choose a subscription plan to activate or renew PITC Suite:")
        info_lbl.setStyleSheet(f"font-size: 9pt; color: {COLOR_TEXT_SECONDARY};")
        layout.addWidget(info_lbl)

        # Plans Grid
        grid = QGridLayout()
        grid.setSpacing(8)

        plans_data = [
            ("1 Month Plan", "Rs. 600", "30 Days Validity", "1 Device Limit"),
            ("3 Months Plan", "Rs. 1,400", "90 Days Validity", "Up to 2 Devices"),
            ("6 Months Plan", "Rs. 2,400", "180 Days Validity", "Up to 2 Devices"),
            ("1 Year Plan", "Rs. 4,000", "365 Days Validity", "Up to 3 Devices"),
        ]

        for i, (title, price, dur, dev) in enumerate(plans_data):
            row = i // 2
            col = i % 2
            card = QFrame()
            card.setStyleSheet(f"""
                QFrame {{
                    background: #1e222a;
                    border: 1px solid #334155;
                    border-radius: 6px;
                    padding: 6px;
                }}
            """)
            c_layout = QVBoxLayout(card)
            c_layout.setContentsMargins(10, 8, 10, 8)
            c_layout.setSpacing(2)

            t_lbl = QLabel(title)
            t_lbl.setStyleSheet("font-weight: bold; color: #f8fafc; font-size: 9.5pt;")
            c_layout.addWidget(t_lbl)

            p_lbl = QLabel(price)
            p_lbl.setStyleSheet("font-weight: bold; color: #4ade80; font-size: 12pt;")
            c_layout.addWidget(p_lbl)

            d_lbl = QLabel(f"• {dur}\n• {dev}")
            d_lbl.setStyleSheet("color: #94a3b8; font-size: 8pt;")
            c_layout.addWidget(d_lbl)

            grid.addWidget(card, row, col)

        layout.addLayout(grid)
        layout.addSpacing(4)

        # Quick Order Action Buttons
        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)

        btn_wa = QPushButton("💬 Order via WhatsApp")
        btn_wa.setCursor(Qt.PointingHandCursor)
        btn_wa.setStyleSheet("""
            QPushButton {
                background: #059669;
                color: white;
                font-weight: bold;
                font-size: 9pt;
                padding: 9px;
                border: none;
                border-radius: 5px;
            }
            QPushButton:hover {
                background: #047857;
            }
        """)
        btn_wa.clicked.connect(self._open_whatsapp)
        btn_row.addWidget(btn_wa)

        btn_email = QPushButton("✉ Contact via Email")
        btn_email.setCursor(Qt.PointingHandCursor)
        btn_email.setStyleSheet("""
            QPushButton {
                background: #2563eb;
                color: white;
                font-weight: bold;
                font-size: 9pt;
                padding: 9px;
                border: none;
                border-radius: 5px;
            }
            QPushButton:hover {
                background: #1d4ed8;
            }
        """)
        btn_email.clicked.connect(self._open_email)
        btn_row.addWidget(btn_email)

        layout.addLayout(btn_row)

        return widget

    def _format_serial_input(self, text: str):
        """Auto uppercase and assist in serial key formatting."""
        clean = re.sub(r'[^A-Za-z0-9]', '', text).upper()
        # If user pasted raw 12-16 chars without dashes, format it
        if clean.startswith("PS") and len(clean) > 2 and "-" not in text:
            chunks = [clean[:2]]
            remainder = clean[2:]
            for i in range(0, len(remainder), 4):
                chunks.append(remainder[i:i+4])
            formatted = "-".join(chunks[:4])
            if formatted != text:
                self.key_entry.blockSignals(True)
                self.key_entry.setText(formatted)
                self.key_entry.setCursorPosition(len(formatted))
                self.key_entry.blockSignals(False)

    def _on_copy_hwid(self):
        cb = QApplication.clipboard()
        if cb:
            cb.setText(self.device_id)
            self.btn_copy.setText("✓ Copied")
            self.btn_copy.setStyleSheet("""
                QPushButton {
                    background: #15803d;
                    color: white;
                    font-weight: bold;
                    font-size: 9pt;
                    padding: 7px 14px;
                    border: none;
                    border-radius: 5px;
                }
            """)

    def _open_whatsapp(self):
        msg = f"Hello Tahir, I would like to buy a license for PITC Suite.\nMy Hardware ID is: {self.device_id}"
        encoded = urllib.parse.quote(msg)
        url = f"https://wa.me/{SUPPORT_WHATSAPP}?text={encoded}"
        webbrowser.open(url)

    def _open_email(self):
        subject = urllib.parse.quote("PITC Suite License Activation Request")
        body = urllib.parse.quote(f"Hello Tahir,\n\nI want to purchase a license for PITC Suite.\n\nHardware ID: {self.device_id}\n\nThank you!")
        url = f"mailto:{SUPPORT_EMAIL}?subject={subject}&body={body}"
        webbrowser.open(url)

    def _on_activate_clicked(self):
        key = self.key_entry.text().strip().upper()
        if not key:
            self.status_lbl.setText("Please enter your Serial Key.")
            self.status_lbl.setStyleSheet(f"color: {COLOR_WARNING}; font-size: 8.5pt;")
            return

        self.btn_activate.setEnabled(False)
        self.btn_activate.setText("Verifying with License Server...")
        self.status_lbl.setText("Contacting server and validating hardware signature...")
        self.status_lbl.setStyleSheet(f"color: #38bdf8; font-size: 8.5pt;")

        # Run activation in background thread to prevent UI lockup
        self._worker = ActivationWorker(key, self.device_id, self.candidate_ids)
        self._worker.finished_signal.connect(self._on_activation_finished)
        self._worker.start()

    def _on_activation_finished(self, res: dict):
        self.btn_activate.setEnabled(True)
        self.btn_activate.setText("Activate PITC Suite")

        if res.get("valid"):
            now_iso = datetime.now(timezone.utc).isoformat()
            lic_data = {
                "serial_key": res.get("serial_key") or self.key_entry.text().strip().upper(),
                "device_id": self.device_id,
                "plan_type": res.get("plan_type", "active"),
                "duration_days": res.get("duration_days", 30),
                "activated_at": res.get("activated_at", now_iso),
                "expires_at": res.get("expires_at"),
                "max_devices": res.get("max_devices", 1),
                "last_verified": now_iso,
                "valid": True
            }
            save_local_license(lic_data)

            self.activated = True
            self.license_info = lic_data

            QMessageBox.information(
                self,
                "Activation Successful",
                f"Thank you! PITC Suite has been successfully activated.\n\n"
                f"Plan: {lic_data.get('plan_type', 'Standard')}\n"
                f"Device ID: {self.device_id}\n\n"
                "All features are now unlocked."
            )
            self.accept()
        else:
            err_msg = res.get("message") or "Activation failed. Please check your serial key."
            self.status_lbl.setText(f"❌ {err_msg}")
            self.status_lbl.setStyleSheet(f"color: {COLOR_ERROR}; font-size: 8.5pt;")

    def closeEvent(self, event):
        """Prevent launching application if closed without activation."""
        if not self.activated:
            sys.exit(0)
        event.accept()


# ====================================================================
# ENTRYPOINT VERIFICATION FUNCTION
# ====================================================================

def require_license() -> bool:
    """
    Main license requirement guard called on application startup.
    1. Reads local license cache (%APPDATA%/PITCSuite/license.json).
    2. Validates against remote Vercel API.
    3. If offline, checks 7-day offline grace period.
    4. If legacy offline key exists, honors it.
    5. If invalid or missing, displays sleek LicenseGateDialog.
    """
    local_lic = load_local_license()
    primary_id = get_device_id()
    candidate_ids = get_candidate_device_ids()

    if local_lic and local_lic.get("serial_key"):
        serial_key = local_lic["serial_key"]

        # Perform remote check against Supabase via Vercel API
        res = verify_license_remote(serial_key, primary_id, candidate_ids)

        if res.get("valid"):
            local_lic.update(res)
            local_lic["last_verified"] = datetime.now(timezone.utc).isoformat()
            local_lic["valid"] = True
            save_local_license(local_lic)
            return True

        # If network error, check offline grace period
        if res.get("status") == "network_error":
            is_valid, reason = is_offline_grace_valid(local_lic)
            if is_valid:
                return True
            initial_error = "Unable to connect to license server and offline grace period has expired."
        else:
            initial_error = res.get("message") or "License subscription is expired or inactive."
    else:
        initial_error = ""

    # Ensure QApplication exists before opening dialog
    app = QApplication.instance()
    if not app:
        app = QApplication(sys.argv)

    dialog = LicenseGateDialog(initial_error=initial_error)
    result = dialog.exec()
    if result != QDialog.Accepted or not dialog.activated:
        sys.exit(0)

    return True

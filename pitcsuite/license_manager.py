import os
import sys
import subprocess
import hashlib
import winreg
import uuid
import re
from typing import List, Tuple, Set

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QMessageBox, QApplication
)
from PySide6.QtCore import Qt

SECRET_SALT = "PITCSUITE_TAHIR_2026_PRO"


def get_hardware_id() -> str:
    """
    Returns the primary hardware ID to display to the user.
    Uses Windows Registry MachineGuid as primary, falling back to CIM/WMIC UUID,
    MAC address hash, or legacy fallback.
    """
    # 1. Windows Registry MachineGuid (64-bit & 32-bit views)
    for flag in [getattr(winreg, 'KEY_WOW64_64KEY', 0), getattr(winreg, 'KEY_WOW64_32KEY', 0), 0]:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Cryptography", 0, winreg.KEY_READ | flag) as key:
                machine_guid, _ = winreg.QueryValueEx(key, "MachineGuid")
                if machine_guid and len(machine_guid.strip()) > 10:
                    return machine_guid.strip().upper()
        except Exception:
            pass

    # 2. Modern PowerShell CIM query for BIOS UUID
    try:
        cmd = ["powershell", "-NoProfile", "-Command", "(Get-CimInstance Win32_ComputerSystemProduct).UUID"]
        creation_flags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
        output = subprocess.check_output(cmd, creationflags=creation_flags, timeout=5).decode().strip()
        if '-' in output and len(output) >= 30:
            return output.upper()
    except Exception:
        pass

    # 3. Legacy WMIC UUID
    try:
        cmd = "wmic csproduct get uuid"
        output = subprocess.check_output(cmd, shell=True, timeout=5).decode().split('\n')
        for line in output:
            if '-' in line and len(line.strip()) >= 30:
                return line.strip().upper()
    except Exception:
        pass

    # 4. Fallback to physical MAC address hash
    try:
        mac = uuid.getnode()
        return f"FALLBACK-{hex(mac)[2:].upper()}"
    except Exception:
        pass

    return "UNKNOWN-HWID-001"


def get_all_machine_hwids() -> List[str]:
    """
    Gathers all legitimate hardware IDs that can identify this specific machine.
    This guarantees full backward and forward compatibility across updates,
    preventing existing activated clients from being locked out.
    """
    hwids: Set[str] = set()

    # 1. Windows Registry MachineGuid (both 64-bit and 32-bit registry views)
    for flag in [getattr(winreg, 'KEY_WOW64_64KEY', 0), getattr(winreg, 'KEY_WOW64_32KEY', 0), 0]:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Cryptography", 0, winreg.KEY_READ | flag) as key:
                val, _ = winreg.QueryValueEx(key, "MachineGuid")
                if val and len(val.strip()) > 10:
                    clean = val.strip()
                    hwids.add(clean)
                    hwids.add(clean.upper())
                    hwids.add(clean.lower())
        except Exception:
            pass

    # 2. PowerShell CIM query (BIOS UUID from Win32_ComputerSystemProduct)
    try:
        cmd = ["powershell", "-NoProfile", "-Command", "(Get-CimInstance Win32_ComputerSystemProduct).UUID"]
        creation_flags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
        out = subprocess.check_output(cmd, creationflags=creation_flags, timeout=5).decode().strip()
        if '-' in out and len(out) >= 30:
            clean = out.strip()
            hwids.add(clean)
            hwids.add(clean.upper())
            hwids.add(clean.lower())
    except Exception:
        pass

    # 3. Legacy WMIC query (BIOS UUID from wmic csproduct)
    try:
        cmd = "wmic csproduct get uuid"
        lines = subprocess.check_output(cmd, shell=True, timeout=5).decode().split('\n')
        for line in lines:
            if '-' in line and len(line.strip()) >= 30:
                clean = line.strip()
                hwids.add(clean)
                hwids.add(clean.upper())
                hwids.add(clean.lower())
    except Exception:
        pass

    # 4. Physical MAC address fallback
    try:
        mac = uuid.getnode()
        hwids.add(f"FALLBACK-{hex(mac)[2:].upper()}")
        hwids.add(f"FALLBACK-{hex(mac)[2:].lower()}")
    except Exception:
        pass

    # 5. Grandfathered legacy fallback identifier
    hwids.add("UNKNOWN-HWID-001")

    return [h for h in hwids if h]


def generate_expected_key(hwid: str) -> str:
    """Generates a 16-character machine-locked serial key formatted as XXXX-XXXX-XXXX-XXXX."""
    if not hwid:
        return ""
    raw = (hwid.strip() + SECRET_SALT).encode('utf-8')
    full_hash = hashlib.sha256(raw).hexdigest().upper()
    return f"{full_hash[:4]}-{full_hash[4:8]}-{full_hash[8:12]}-{full_hash[12:16]}"


def normalize_key(key: str) -> str:
    """Removes hyphens, spaces, and normalizes to uppercase 16 hex characters."""
    if not key:
        return ""
    return re.sub(r'[^A-Fa-f0-9]', '', key).upper()


def is_key_valid_for_machine(candidate_key: str) -> Tuple[bool, str, str]:
    """
    Checks if a candidate serial key is valid for this machine under ANY of its valid HWIDs.
    Returns: (is_valid, matched_hwid, expected_formatted_key)
    """
    norm_candidate = normalize_key(candidate_key)
    if len(norm_candidate) != 16:
        return False, "", ""

    for hwid in get_all_machine_hwids():
        expected = generate_expected_key(hwid)
        if normalize_key(expected) == norm_candidate:
            return True, hwid, expected

    return False, "", ""


def get_license_search_paths() -> List[str]:
    """Returns all redundant persistent storage paths for the license key."""
    paths = []

    # 1. Roaming AppData (User profile)
    appdata = os.getenv('APPDATA') or os.path.expanduser(r'~\AppData\Roaming')
    if appdata:
        paths.append(os.path.join(appdata, 'PITCSuite', 'license.key'))

    # 2. Local AppData (Installation & local caches)
    localappdata = os.getenv('LOCALAPPDATA') or os.path.expanduser(r'~\AppData\Local')
    if localappdata:
        paths.append(os.path.join(localappdata, 'PITCSuite', 'license.key'))
        paths.append(os.path.join(localappdata, 'Programs', 'PITCSuite', 'license.key'))

    # 3. Machine-wide ProgramData (Shared by all users & elevated installer/updater runs)
    programdata = os.getenv('PROGRAMDATA') or os.getenv('ALLUSERSPROFILE') or r'C:\ProgramData'
    if programdata:
        paths.append(os.path.join(programdata, 'PITCSuite', 'license.key'))

    # 4. User profile home directory
    userprofile = os.getenv('USERPROFILE') or os.path.expanduser('~')
    if userprofile:
        paths.append(os.path.join(userprofile, '.pitcsuite', 'license.key'))

    # 5. Application directory (if portable / writable)
    try:
        if getattr(sys, 'frozen', False):
            app_dir = os.path.dirname(sys.executable)
        else:
            app_dir = os.path.dirname(os.path.abspath(__file__))
        paths.append(os.path.join(app_dir, 'license.key'))
    except Exception:
        pass

    seen = set()
    deduped = []
    for p in paths:
        norm = os.path.normpath(p).lower()
        if norm not in seen:
            seen.add(norm)
            deduped.append(os.path.normpath(p))
    return deduped


def get_license_file_path() -> str:
    """Primary license file path for backward compatibility."""
    appdata = os.getenv('APPDATA') or os.path.expanduser(r'~\AppData\Roaming')
    folder = os.path.join(appdata, 'PITCSuite')
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, 'license.key')


def load_saved_license_keys() -> List[str]:
    """Retrieves all candidate license keys stored across persistent file locations and Windows Registry."""
    found_keys: List[str] = []

    # 1. Search persistent file paths
    for path in get_license_search_paths():
        if os.path.exists(path):
            try:
                with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                    for line in f:
                        # Extract any XXXX-XXXX-XXXX-XXXX
                        match = re.search(r'[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}', line)
                        if match:
                            found_keys.append(match.group(0).upper())
                        else:
                            clean_line = normalize_key(line.strip())
                            if len(clean_line) == 16:
                                formatted = f"{clean_line[:4]}-{clean_line[4:8]}-{clean_line[8:12]}-{clean_line[12:16]}"
                                found_keys.append(formatted)
            except Exception:
                pass

    # 2. Search HKCU Registry
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\PITCSuite", 0, winreg.KEY_READ) as k:
            val, _ = winreg.QueryValueEx(k, "LicenseKey")
            if val:
                match = re.search(r'[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}', str(val))
                if match:
                    found_keys.append(match.group(0).upper())
    except Exception:
        pass

    # 3. Search HKLM Registry (64-bit and 32-bit views)
    for flag in [getattr(winreg, 'KEY_WOW64_64KEY', 0), getattr(winreg, 'KEY_WOW64_32KEY', 0), 0]:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"Software\PITCSuite", 0, winreg.KEY_READ | flag) as k:
                val, _ = winreg.QueryValueEx(k, "LicenseKey")
                if val:
                    match = re.search(r'[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}', str(val))
                    if match:
                        found_keys.append(match.group(0).upper())
        except Exception:
            pass

    seen = set()
    deduped = []
    for k in found_keys:
        if k not in seen:
            seen.add(k)
            deduped.append(k)
    return deduped


def save_license_key(keys_to_save) -> None:
    """Saves valid license key(s) to all persistent user, system files, and Windows Registry."""
    if isinstance(keys_to_save, str):
        keys_list = [keys_to_save]
    else:
        keys_list = list(keys_to_save)

    clean_keys: List[str] = []
    for k in keys_list:
        if not k:
            continue
        norm = normalize_key(k)
        if len(norm) == 16:
            formatted = f"{norm[:4]}-{norm[4:8]}-{norm[8:12]}-{norm[12:16]}"
            if formatted not in clean_keys:
                clean_keys.append(formatted)

    if not clean_keys:
        return

    file_content = "\n".join(clean_keys) + "\n"

    # 1. Write to all persistent file paths
    for path in get_license_search_paths():
        try:
            folder = os.path.dirname(path)
            os.makedirs(folder, exist_ok=True)
            with open(path, 'w', encoding='utf-8') as f:
                f.write(file_content)
        except Exception:
            pass

    # 2. Write to HKCU Registry
    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\PITCSuite") as k:
            winreg.SetValueEx(k, "LicenseKey", 0, winreg.REG_SZ, clean_keys[0])
    except Exception:
        pass

    # 3. Write to HKLM Registry (if administrator privileges permit)
    try:
        with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, r"Software\PITCSuite") as k:
            winreg.SetValueEx(k, "LicenseKey", 0, winreg.REG_SZ, clean_keys[0])
    except Exception:
        pass


class LicenseDialog(QDialog):
    def __init__(self, hwid: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("PITC Suite - Product Activation")
        self.setFixedSize(460, 280)
        self.setWindowFlags(Qt.Dialog | Qt.CustomizeWindowHint | Qt.WindowTitleHint)  # No close button

        self.setStyleSheet("""
            QDialog {
                background: #1e222a;
                color: #edf0f5;
                font-family: 'Segoe UI', Tahoma, sans-serif;
            }
            QLabel {
                color: #cbd5e1;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(10)

        lbl1 = QLabel("This software requires a serial key to run.")
        lbl1.setStyleSheet("font-weight: bold; font-size: 11pt; color: #f8fafc;")
        layout.addWidget(lbl1)

        lbl2 = QLabel("Please send the following Hardware ID to tahir.naqash4@gmail.com or contact +923339962474:")
        lbl2.setWordWrap(True)
        lbl2.setStyleSheet("font-size: 9pt; color: #94a3b8;")
        layout.addWidget(lbl2)

        # HWID row with Copy button
        hwid_row = QHBoxLayout()
        hwid_row.setSpacing(8)

        self.hwid_input = QLineEdit(hwid)
        self.hwid_input.setReadOnly(True)
        self.hwid_input.setStyleSheet("""
            QLineEdit {
                background: #0f172a;
                color: #38bdf8;
                font-weight: bold;
                font-size: 11pt;
                padding: 6px 10px;
                border: 1px solid #334155;
                border-radius: 5px;
            }
        """)
        self.hwid_input.setAlignment(Qt.AlignCenter)
        hwid_row.addWidget(self.hwid_input, 1)

        self.copy_btn = QPushButton("📋 Copy")
        self.copy_btn.setCursor(Qt.PointingHandCursor)
        self.copy_btn.setStyleSheet("""
            QPushButton {
                background: #334155;
                color: #f1f5f9;
                font-weight: bold;
                font-size: 9pt;
                padding: 6px 14px;
                border: none;
                border-radius: 5px;
            }
            QPushButton:hover {
                background: #475569;
            }
        """)
        self.copy_btn.clicked.connect(self.copy_hwid)
        hwid_row.addWidget(self.copy_btn)

        layout.addLayout(hwid_row)

        lbl3 = QLabel("Enter your Serial Key:")
        lbl3.setStyleSheet("font-size: 9.5pt; font-weight: bold; color: #e2e8f0; margin-top: 4px;")
        layout.addWidget(lbl3)

        self.key_input = QLineEdit()
        self.key_input.setPlaceholderText("XXXX-XXXX-XXXX-XXXX")
        self.key_input.setStyleSheet("""
            QLineEdit {
                background: #0f172a;
                color: #4ade80;
                font-weight: bold;
                font-size: 12pt;
                padding: 6px;
                border: 1px solid #334155;
                border-radius: 5px;
            }
        """)
        self.key_input.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.key_input)

        self.activate_btn = QPushButton("Activate Software")
        self.activate_btn.setCursor(Qt.PointingHandCursor)
        self.activate_btn.setStyleSheet("""
            QPushButton {
                background: #16a34a;
                color: white;
                font-weight: bold;
                font-size: 10pt;
                padding: 9px;
                border: none;
                border-radius: 6px;
                margin-top: 4px;
            }
            QPushButton:hover {
                background: #15803d;
            }
        """)
        self.activate_btn.clicked.connect(self.check_activation)
        layout.addWidget(self.activate_btn)

        self.hwid = hwid

    def copy_hwid(self):
        clipboard = QApplication.clipboard()
        if clipboard:
            clipboard.setText(self.hwid)
            self.copy_btn.setText("✓ Copied")
            self.copy_btn.setStyleSheet("""
                QPushButton {
                    background: #15803d;
                    color: white;
                    font-weight: bold;
                    font-size: 9pt;
                    padding: 6px 14px;
                    border: none;
                    border-radius: 5px;
                }
            """)

    def check_activation(self):
        user_key = self.key_input.text().strip()
        is_valid, matched_hwid, expected_key = is_key_valid_for_machine(user_key)

        if is_valid:
            primary_hwid = get_hardware_id()
            primary_key = generate_expected_key(primary_hwid)
            # Save both the primary key and the entered valid key across all locations
            save_license_key([primary_key, expected_key, user_key])
            QMessageBox.information(self, "Activated", "Thank you! Software activated successfully.")
            self.accept()
        else:
            QMessageBox.critical(self, "Error", "Invalid Serial Key! Please check and try again.")


def require_license() -> bool:
    """
    Checks if a valid license exists across any storage location for any valid machine HWID.
    If valid, synchronizes across all storage tiers and returns True without prompting.
    If no valid license exists, prompts with LicenseDialog.
    """
    saved_keys = load_saved_license_keys()

    for key in saved_keys:
        is_valid, matched_hwid, expected_key = is_key_valid_for_machine(key)
        if is_valid:
            # Silently synchronize & ensure primary key and matched key are preserved everywhere
            primary_hwid = get_hardware_id()
            primary_key = generate_expected_key(primary_hwid)
            save_license_key([primary_key, expected_key, key])
            return True

    # If no valid license, ensure QApplication exists and show dialog
    app = QApplication.instance()
    if not app:
        app = QApplication(sys.argv)

    primary_hwid = get_hardware_id()
    dialog = LicenseDialog(primary_hwid)
    result = dialog.exec()
    if result != QDialog.Accepted:
        sys.exit(0)

    return True

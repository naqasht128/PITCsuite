import os
import sys
import subprocess
import hashlib
import winreg
import uuid
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QLineEdit, QPushButton, QMessageBox, QApplication
)
from PySide6.QtCore import Qt

SECRET_SALT = "PITCSUITE_TAHIR_2026_PRO"

def get_hardware_id():
    """Gets a unique machine ID (UUID) for Windows reliably across all versions."""
    # Method 1: Read MachineGuid from Windows Registry (Fastest & 100% reliable on Win 7/8/10/11)
    try:
        registry_path = r"SOFTWARE\Microsoft\Cryptography"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, registry_path, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as key:
            machine_guid, _ = winreg.QueryValueEx(key, "MachineGuid")
            if machine_guid and len(machine_guid.strip()) > 10:
                return machine_guid.strip().upper()
    except Exception:
        pass

    # Method 2: Modern PowerShell CIM query (works on Win 10/11 even if WMIC is removed)
    try:
        cmd = ["powershell", "-NoProfile", "-Command", "(Get-CimInstance Win32_ComputerSystemProduct).UUID"]
        creation_flags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
        output = subprocess.check_output(cmd, creationflags=creation_flags).decode().strip()
        if '-' in output and len(output) >= 30:
            return output.upper()
    except Exception:
        pass

    # Method 3: Legacy wmic (for older Windows versions)
    try:
        cmd = "wmic csproduct get uuid"
        output = subprocess.check_output(cmd, shell=True).decode().split('\n')
        for line in output:
            if '-' in line:
                return line.strip().upper()
    except Exception:
        pass

    # Method 4: Fallback to physical MAC address hash
    try:
        mac = uuid.getnode()
        return f"FALLBACK-{hex(mac)[2:].upper()}"
    except Exception:
        return "UNKNOWN-HWID-001"

def generate_expected_key(hwid):
    raw = (hwid.strip() + SECRET_SALT).encode('utf-8')
    full_hash = hashlib.sha256(raw).hexdigest().upper()
    return f"{full_hash[:4]}-{full_hash[4:8]}-{full_hash[8:12]}-{full_hash[12:16]}"

def get_license_file_path():
    appdata = os.getenv('APPDATA')
    folder = os.path.join(appdata, 'PITCSuite')
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, 'license.key')

class LicenseDialog(QDialog):
    def __init__(self, hwid, parent=None):
        super().__init__(parent)
        self.setWindowTitle("PITC Suite - Product Activation")
        self.setFixedSize(400, 250)
        self.setWindowFlags(Qt.Dialog | Qt.CustomizeWindowHint | Qt.WindowTitleHint) # No close button
        
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        
        lbl1 = QLabel("This software requires a serial key to run.")
        lbl1.setStyleSheet("font-weight: bold; font-size: 10pt;")
        layout.addWidget(lbl1)
        
        lbl2 = QLabel("Please email the following Hardware ID to tahir.naqash4@gmail.com or contact +923339962474 for your serial key:")
        layout.addWidget(lbl2)
        
        self.hwid_input = QLineEdit(hwid)
        self.hwid_input.setReadOnly(True)
        self.hwid_input.setStyleSheet("background: #1e293b; color: #38bdf8; font-weight: bold; font-size: 12pt;")
        self.hwid_input.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.hwid_input)
        
        lbl3 = QLabel("Enter your Serial Key:")
        layout.addWidget(lbl3)
        
        self.key_input = QLineEdit()
        self.key_input.setPlaceholderText("XXXX-XXXX-XXXX-XXXX")
        self.key_input.setStyleSheet("font-size: 11pt;")
        self.key_input.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.key_input)
        
        self.activate_btn = QPushButton("Activate Software")
        self.activate_btn.setStyleSheet("background: #059669; color: white; font-weight: bold; padding: 8px;")
        self.activate_btn.clicked.connect(self.check_activation)
        layout.addWidget(self.activate_btn)
        
        self.hwid = hwid

    def check_activation(self):
        user_key = self.key_input.text().strip().upper()
        expected_key = generate_expected_key(self.hwid)
        
        if user_key == expected_key:
            QMessageBox.information(self, "Activated", "Thank you! Software activated successfully.")
            with open(get_license_file_path(), 'w') as f:
                f.write(user_key)
            self.accept()
        else:
            QMessageBox.critical(self, "Error", "Invalid Serial Key! Please check and try again.")

def require_license():
    """Checks if valid license exists. If not, blocks until activated."""
    hwid = get_hardware_id()
    expected_key = generate_expected_key(hwid)
    
    license_file = get_license_file_path()
    if os.path.exists(license_file):
        try:
            with open(license_file, 'r') as f:
                saved_key = f.read().strip()

            # 1. Valid against current machine's real HWID
            if saved_key == expected_key:
                return True # Valid

            # 2. Backward compatibility: grandfather existing client who had UNKNOWN-HWID-001
            legacy_key = generate_expected_key("UNKNOWN-HWID-001")
            if saved_key == legacy_key:
                # Silently upgrade their saved license file to their new genuine HWID key
                try:
                    with open(license_file, 'w') as fw:
                        fw.write(expected_key)
                except Exception:
                    pass
                return True # Validated & upgraded seamlessly
        except Exception:
            pass
                
    # If no valid license, show dialog
    # We create a dummy app if it doesn't exist yet, but in our case QApplication is already running
    dialog = LicenseDialog(hwid)
    result = dialog.exec()
    if result != QDialog.Accepted:
        sys.exit(0) # Exit if they cancel/close it somehow



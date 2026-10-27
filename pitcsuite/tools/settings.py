import os
import re
import time
import base64
import threading
import configparser

from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton,
    QLineEdit, QFileDialog, QCheckBox, QTextEdit, QProgressBar,
    QVBoxLayout, QHBoxLayout, QGroupBox, QMessageBox, QComboBox,
    QSpinBox
)
from PySide6.QtCore import Signal, QObject, QThread, QSettings

from pitcsuite.ui_helpers import lbl, hline, START_BTN, STOP_BTN
from pitcsuite.templates import download_template
from pitcsuite.config import config_path as _config_path

class SettingsPanel(QWidget):
    DISCOS = [
        "PESCO", "LESCO", "GEPCO", "FESCO", "IESCO", "MEPCO",
        "HAZECO", "SEPCO", "HESCO", "QESCO", "TESCO"
    ]

    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self); layout.setSpacing(14)

        layout.addWidget(lbl("Settings  ·  PITC Credentials", bold=True, color="#7eb8f7"))
        layout.addWidget(hline())

        g = QGroupBox("PITC Login (cp.pitc.com.pk)")
        gv = QVBoxLayout(g); gv.setSpacing(10)

        r0 = QHBoxLayout()
        r0.addWidget(QLabel("Default DISCO:")); self.disco_cb = QComboBox(); self.disco_cb.addItems(self.DISCOS)
        self.disco_cb.setCurrentText("PESCO"); r0.addWidget(self.disco_cb, 1)
        gv.addLayout(r0)

        r1 = QHBoxLayout()
        r1.addWidget(QLabel("Username:")); self.user_ed = QLineEdit(); r1.addWidget(self.user_ed, 1)
        gv.addLayout(r1)

        r2 = QHBoxLayout()
        r2.addWidget(QLabel("Password:")); self.pass_ed = QLineEdit()
        self.pass_ed.setEchoMode(QLineEdit.Password); r2.addWidget(self.pass_ed, 1)
        self.chk_show = QCheckBox("Show password")
        self.chk_show.toggled.connect(lambda on: self.pass_ed.setEchoMode(QLineEdit.Normal if on else QLineEdit.Password))
        r2.addWidget(self.chk_show)
        gv.addLayout(r2)

        r3 = QHBoxLayout()
        r3.addWidget(QLabel("OTP Secret:")); self.otp_ed = QLineEdit(); self.otp_ed.setProperty("preferred_width", 320)
        self.otp_ed.setEchoMode(QLineEdit.Password)
        self.otp_ed.setPlaceholderText("TOTP base32 secret (leave blank if no OTP required)")
        r3.addWidget(self.otp_ed, 1)
        self.chk_otp_show = QCheckBox("Show")
        self.chk_otp_show.toggled.connect(lambda on: self.otp_ed.setEchoMode(QLineEdit.Normal if on else QLineEdit.Password))
        r3.addWidget(self.chk_otp_show)
        gv.addLayout(r3)

        layout.addWidget(g)

        note = QLabel(f"Saved to:  {_config_path()}")
        note.setStyleSheet("color:#556; font-size:10px;")
        note.setWordWrap(True)
        layout.addWidget(note)
        self.path_lbl = note  # update after save

        bh = QHBoxLayout()
        btn_save = QPushButton("💾  Save Credentials")
        btn_save.setProperty("preferred_width", 170)
        btn_save.setStyleSheet(START_BTN + " padding:8px 24px;")
        btn_save.clicked.connect(self.save)
        bh.addWidget(btn_save); bh.addStretch()
        layout.addLayout(bh)

        self.status_lbl = QLabel("")
        self.status_lbl.setStyleSheet("color:#4c4; font-size:11px;")
        layout.addWidget(self.status_lbl)

        layout.addStretch()
        layout.addWidget(lbl("Tip: credentials are stored in plain text in config.ini.\nKeep the file in a safe location.", color="#556"))

        # load existing values on open
        self.load()

    def load(self):
        saved_disco = "PESCO"
        try:
            settings = QSettings("TNDevelopers", "PITCSuite")
            if settings.contains("pitc_disco"):
                saved_disco = str(settings.value("pitc_disco")).strip() or "PESCO"
        except Exception:
            pass

        path = _config_path()
        cfg = configparser.ConfigParser()
        if os.path.exists(path):
            cfg.read(path)
            self.user_ed.setText(cfg.get("credentials", "username", fallback=""))
            self.pass_ed.setText(cfg.get("credentials", "password", fallback=""))
            self.otp_ed.setText(cfg.get("credentials", "otp_secret", fallback=""))
            saved_disco = cfg.get("credentials", "disco", fallback=cfg.get("pitc_downloader", "disco", fallback=saved_disco))
            self.status_lbl.setText("")
        else:
            self.status_lbl.setText("⚠  config.ini not found — will be created on Save.")
            self.status_lbl.setStyleSheet("color:#ca4; font-size:11px;")

        idx = self.disco_cb.findText(saved_disco)
        if idx >= 0:
            self.disco_cb.setCurrentIndex(idx)

    def save(self):
        username   = self.user_ed.text().strip()
        password   = self.pass_ed.text().strip()
        otp_secret = self.otp_ed.text().strip()
        disco      = self.disco_cb.currentText().strip() or "PESCO"
        if not username or not password:
            QMessageBox.warning(self, "Missing", "Username and password cannot be empty."); return

        try:
            settings = QSettings("TNDevelopers", "PITCSuite")
            settings.setValue("pitc_disco", disco)
        except Exception:
            pass

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

        self.status_lbl.setText("✔  Saved successfully.")
        self.status_lbl.setStyleSheet("color:#4c4; font-size:11px;")
        self.path_lbl.setText(f"Saved to:  {path}")


# ──────────────────────────────────────────────────────────────
#  TOOL 10 — Extract Specific PITC Page  (ExtPITC.py)
#  Search Criteria: month (Feb-26 → Feb - 2026) or any number
#  Extracts matching page · Highlights matches · Stamps Sr. No.
# ──────────────────────────────────────────────────────────────

"""
Auto-updater module for PITC Suite.
Handles version checks, download progress, silent/manual updates,
and displays comprehensive "What's New" release notes / amendment bullets.
"""

import os
import sys
import subprocess
import requests
import re
import json
import threading
import time
from typing import List, Dict, Tuple, Optional
from packaging import version
from PySide6.QtWidgets import (
    QMessageBox, QProgressDialog, QApplication, QDialog,
    QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QFrame,
    QScrollArea, QWidget
)
from PySide6.QtCore import QThread, Signal, Qt, QObject, QTimer, QSettings
from PySide6.QtGui import QFont, QCursor

GDRIVE_VERSION_FILE_ID = "1gmIKNgsj2hdYrrQtiwXHyQcXzCDR-ZZb" 
GDRIVE_SETUP_FILE_ID = "1wghUxDo3ezed5szbpXeWtWlyof4ykBIu"

# Comprehensive 2-3 line release notes detailing recent amendments
DEFAULT_RELEASE_NOTES: Dict[str, List[str]] = {
    "v2.8.91": [
        "<b>Brand New Tool — Govt Assessment Realization:</b> Native tool to import WAPDA Assessment & Realization text reports (.txt) and extract structured consumer billing data into professionally styled Excel workbooks matching official templates.",
        "<b>Exact 12-Column Extraction:</b> Accurately extracts Ref No, Batch No, Sub Div, Cons Name, Month, Triff, Reading, Assessed Amount, Total Amount, Realized Amount, Balance Amount, and Status (Disconnected/Running).",
        "<b>Streamlined Single-File UI:</b> Simplified single-file import with automatic Excel output naming, fast processing, and one-click opening of output files and folders.",
        "<b>UI Refinements & Renaming:</b> Renamed 'CP136_137 Posting Checker' to 'CP 136 & 137 Posting Checker' and streamlined CP 139 Posting Checker interface."
    ],
    "v2.8.90": [
        "<b>CP 139 Proof List & Adjustor Support:</b> Added complete parsing support for WAPDA CP 139 Proof List (Adjustor) reports, extracting record numbers, 13-digit and 14-digit reference numbers, and auto-deriving 5-digit sub-division codes.",
        "<b>RFP & NR Status & Error Isolation:</b> Accurately maps <code>RFP</code> (Ready For Posting) to <i>Posted</i> and <code>NR</code> (Not Recorded) to <i>Not Pos</i>, extracting error reason codes (e.g. <code>R27,R23,NPV</code>, <code>NPV</code>, <code>DUP</code>, <code>G09</code>).",
        "<b>Resilient File Handling & Reporting:</b> Automatically resolves file paths with or without file extensions, generates formatted Excel reports matching report statistical summaries with 100% precision, and supports single-posting deduplication."
    ],
    "v2.8.89": [
        "<b>Brand New Tool — CP 139 Posting Checker:</b> Native integration for parsing WAPDA RCP-139 Unidentified Cash / Debtor's Adjustment posting reports (.txt and .pdf), extracting 14-digit Ref#, 5-digit Sub-Division code, Grand Total amounts, and posting status/errors.",
        "<b>Direct & Streamlined Processing:</b> Export directly to formatted, styled Excel reports with live statistics summary cards without preview table overhead for maximum speed.",
        "<b>Horizontal Scrolling:</b> Enabled smooth horizontal trackpad, tilt-wheel, and shift+wheel scrolling across the entire right pane for large screens and data tables.",
        "<b>Tool Renaming:</b> Renamed 'CP 52 Posting Checker' to 'CP 52 Posting to PDF' to better reflect its highlighted PDF generation workflow.",
        "<b>Automated Update Experience:</b> Visible installation progress bar without requiring manual 'Next' or 'Finish' clicks, automatically restarting upon completion."
    ],
    "v2.8.88": [
        "<b>Brand New Tool — CP 139 Posting Checker:</b> Native integration for parsing WAPDA RCP-139 Unidentified Cash / Debtor's Adjustment posting reports (.txt and .pdf), extracting 14-digit Ref#, 5-digit Sub-Division code, Grand Total amounts, and posting status/errors.",
        "<b>Accurate Parsing & Multi-File Support:</b> Accurately extracts multi-line record layouts, isolating error reason codes (e.g., G09 for unposted records) and handling recursive folder scans or batch multi-file selections.",
        "<b>Deduplication & Status Filters:</b> Advanced single postings deduplication engine with configurable priority strategies (<i>Prioritize POSTED</i>, <i>Keep 1st Finding</i>, <i>Keep Last Finding</i>) and instant filtering for Posted vs. Not Posted transactions.",
        "<b>Live Analytics & Styled Excel Reports:</b> Interactive preview table with real-time statistics cards (Total Records, Posted, Not Pos, Grand Total Sum in Rs.) and professional Excel export with custom header styling, numerical formatting, and color-coded status rows.",
        "<b>Tool Renaming:</b> Renamed 'CP 52 Posting Checker' to 'CP 52 Posting to PDF' to better reflect its highlighted PDF generation workflow."
    ],
    "v2.8.87": [
        "<b>Seamless Auto-Update & Process Termination:</b> Hardened application termination upon update to immediately release file locks and prevent 'file in use' errors.",
        "<b>Automated Visible Installation:</b> Automated installer execution with active progress bar display—no manual 'Next' clicks required—and seamless relaunch upon completion.",
        "<b>Multi-Tier Process Safety:</b> Inno Setup multi-tier process termination ensures all lingering background processes are cleared before files are updated."
    ],
    "v2.8.86": [
        "<b>CP-136 & CP-137 Report Parser Precision:</b> Exact horizontal extraction against headings for middle blocks (Consumer Name, Address, Contact, Net Metering attributes, License#) and column-aligned below-heading extraction for top, Block 1, Block 2, and meter tables.",
        "<b>Error Codes & Meter Reading Protection:</b> Eliminated column cross-contamination, ensuring error codes (A11, TR02, etc.) and posting statuses never leak into S-Points, Meter Reading, or Email Address.",
        "<b>Installer Modernization:</b> Configured Inno Setup 6 64-bit mode (<code>x64compatible</code>) for seamless, robust installation on modern Windows 10 and Windows 11."
    ],
    "v2.8.85": [
        "<b>License Manager Multi-Tier Persistence:</b> Enhanced hardware-locked licensing with seamless backward compatibility across previous BIOS/WMIC UUIDs and modern Registry MachineGuids.",
        "<b>Multi-Location Redundancy:</b> Added automatic license sync across AppData, LocalAppData, ProgramData, and Windows Registry so updates never prompt for serial key again.",
        "<b>Installer Modernization:</b> Updated Inno Setup 6 architecture to <code>x64compatible</code> for full compatibility with both x64 and modern ARM64 Windows."
    ],
    "v2.8.84": [
        "<b>CP136 & 137 Posting Checker:</b> Full support for CP-137 change reports alongside CP-136, automatically generating segregated sheets (<i>Summary</i>, <i>CP136 Posting Data</i>, and <i>CP137 Posting Data</i>) in combined Excel exports.",
        "<b>Sub Division Calculation:</b> Added automated <code>=MID(Ref No., 3, 5)</code> formula column next to Reference Number for instant subdivision grouping and filtering.",
        "<b>High-Precision PDF & TXT Parser:</b> Replaced coordinate clipping with geometric token extraction, preserving 10-digit meter numbers, 14-digit reference alterations, and Net Metering attributes without data loss."
    ],
    "v2.8.83": [
        "<b>Hardware-Locked License Manager:</b> Integrated machine hardware-locked licensing and activation dialog for secure access.",
        "<b>Performance & Stability:</b> Background thread optimization and general reliability improvements across tools."
    ]
}


def get_current_version() -> str:
    """Returns the current application version string."""
    try:
        import pitcsuite
        return f"v{pitcsuite.__version__.lstrip('vV')}"
    except Exception:
        return "v2.8.91"


def get_release_notes(ver: str) -> List[str]:
    """Retrieve comprehensive release notes bullets for a given version tag."""
    clean_ver = ver.strip()
    if clean_ver in DEFAULT_RELEASE_NOTES:
        return DEFAULT_RELEASE_NOTES[clean_ver]
    with_v = f"v{clean_ver.lstrip('vV')}"
    if with_v in DEFAULT_RELEASE_NOTES:
        return DEFAULT_RELEASE_NOTES[with_v]
    return DEFAULT_RELEASE_NOTES.get("v2.8.91", [])


def parse_version_response(raw_text: str) -> Tuple[str, List[str]]:
    """
    Parses version response text from Google Drive or remote URL.
    Supports:
      1. Plain text with version on line 1, and optional notes/bullets on subsequent lines
      2. JSON format: {"version": "v2.8.86", "notes": ["...", "..."]}
    """
    raw_text = raw_text.strip()
    if not raw_text:
        return "", []

    # Attempt JSON parsing
    if (raw_text.startswith("{") and raw_text.endswith("}")) or (raw_text.startswith("[") and raw_text.endswith("]")):
        try:
            data = json.loads(raw_text)
            if isinstance(data, dict):
                ver = str(data.get("version", "")).strip()
                notes = data.get("notes", [])
                if isinstance(notes, str):
                    notes = [n.strip() for n in notes.splitlines() if n.strip()]
                elif isinstance(notes, list):
                    notes = [str(n).strip() for n in notes if str(n).strip()]
                return ver, notes
        except Exception:
            pass

    # Multiline text parsing
    lines = [line.strip() for line in raw_text.splitlines() if line.strip()]
    if not lines:
        return "", []

    first_line = lines[0]
    ver_match = re.search(r'v?\d+(\.\d+)+', first_line, re.IGNORECASE)
    ver = ver_match.group(0) if ver_match else first_line
    if not ver.lower().startswith('v'):
        ver = f"v{ver}"

    notes = []
    for line in lines[1:]:
        cleaned = re.sub(r'^[•\-\*\d\.\)\s]+', '', line).strip()
        if cleaned:
            notes.append(cleaned)

    return ver, notes


def get_gdrive_url_and_token(file_id: str) -> str:
    url = "https://docs.google.com/uc?export=download"
    session = requests.Session()
    response = session.get(url, params={'id': file_id}, stream=True)
    token = None
    for key, value in response.cookies.items():
        if key.startswith('download_warning'):
            token = value
            break
    if not token:
        match = re.search(r'confirm=([a-zA-Z0-9_-]+)', response.text)
        if match:
            token = match.group(1)
            
    final_url = f"{url}&id={file_id}"
    if token:
        final_url += f"&confirm={token}"
    return final_url


class DownloadThread(QThread):
    progress = Signal(int, int)
    finished = Signal(str)
    error = Signal(str)

    def __init__(self, target: str, dest: str):
        super().__init__()
        self.target = target
        self.dest = dest

    def run(self):
        try:
            # 1. Direct or Vercel Proxy streaming URL (HTTP/HTTPS)
            if self.target and (self.target.startswith("http://") or self.target.startswith("https://")) and "docs.google.com" not in self.target and "drive.google.com" not in self.target:
                session = requests.Session()
                response = session.get(self.target, stream=True, timeout=30)
                response.raise_for_status()
                total_size = int(response.headers.get('content-length', 0))
                downloaded = 0
                with open(self.dest, "wb") as f:
                    for chunk in response.iter_content(chunk_size=65536):
                        if chunk:
                            f.write(chunk)
                            downloaded += len(chunk)
                            self.progress.emit(downloaded, total_size)
                self.finished.emit(self.dest)
                return

            # 2. Google Drive File ID / URL fallback
            file_id = self.target if self.target and not self.target.startswith("http") else GDRIVE_SETUP_FILE_ID
            url = "https://docs.google.com/uc?export=download"
            session = requests.Session()
            response = session.get(url, params={'id': file_id}, stream=True)
            
            if "text/html" in response.headers.get("Content-Type", ""):
                text = response.text
                match = re.search(r'name="uuid" value="([^"]+)"', text)
                if match:
                    uuid_val = match.group(1)
                    final_url = f"https://drive.usercontent.google.com/download?id={file_id}&export=download&confirm=t&uuid={uuid_val}"
                    response = session.get(final_url, stream=True)
                else:
                    token = None
                    for key, value in response.cookies.items():
                        if key.startswith('download_warning'):
                            token = value
                            break
                    if not token:
                        match2 = re.search(r'confirm=([a-zA-Z0-9_-]+)', text)
                        if match2:
                            token = match2.group(1)
                    final_url = f"{url}&id={file_id}"
                    if token:
                        final_url += f"&confirm={token}"
                    response = session.get(final_url, stream=True)

            response.raise_for_status()
            total_size = int(response.headers.get('content-length', 0))
            
            downloaded = 0
            with open(self.dest, "wb") as file:
                for chunk in response.iter_content(chunk_size=16384):
                    if chunk:
                        file.write(chunk)
                        downloaded += len(chunk)
                        self.progress.emit(downloaded, total_size)
            self.finished.emit(self.dest)
        except Exception as e:
            self.error.emit(str(e))


class UpdateChecker(QThread):
    result = Signal(bool, str, list, str)  # update_available, version_string, release_notes, download_url

    def run(self):
        # 1. Primary: Check secure Vercel License/Update API
        try:
            from pitcsuite.license_manager import get_api_base_url
            api_url = get_api_base_url()
            latest_url = f"{api_url}/api/updates/latest"
            resp = requests.get(latest_url, timeout=7)
            if resp.status_code == 200:
                data = resp.json()
                raw_ver = str(data.get("version", "")).strip()
                if raw_ver:
                    clean_tag = f"v{raw_ver.lstrip('vV')}"
                    download_url = str(data.get("download_url", "")).strip()
                    server_notes = data.get("release_notes", "")

                    notes_list = []
                    if isinstance(server_notes, list):
                        notes_list = [str(n).strip() for n in server_notes if str(n).strip()]
                    elif isinstance(server_notes, str) and server_notes.strip():
                        notes_list = [n.strip() for n in server_notes.splitlines() if n.strip()]

                    if not notes_list:
                        notes_list = get_release_notes(clean_tag)

                    v_latest = version.parse(clean_tag.lstrip('vV'))
                    v_current = version.parse(get_current_version().lstrip('vV'))
                    if v_latest > v_current:
                        self.result.emit(True, clean_tag, notes_list, download_url)
                        return
                    else:
                        self.result.emit(False, clean_tag, notes_list, download_url)
                        return
        except Exception:
            pass

        # 2. Secondary Fallback: Check Google Drive version file
        try:
            version_url = get_gdrive_url_and_token(GDRIVE_VERSION_FILE_ID)
            v_response = requests.get(version_url, timeout=8)
            v_response.raise_for_status()
            raw_text = v_response.text.strip()
            
            latest_tag, remote_notes = parse_version_response(raw_text)
            
            if latest_tag and latest_tag.lower().startswith("v"):
                v_latest = version.parse(latest_tag.lstrip('vV'))
                v_current = version.parse(get_current_version().lstrip('vV'))
                notes = remote_notes if remote_notes else get_release_notes(latest_tag)
                if v_latest > v_current:
                    self.result.emit(True, latest_tag, notes, GDRIVE_SETUP_FILE_ID)
                    return
                else:
                    self.result.emit(False, latest_tag, notes, "")
                    return
            notes = remote_notes if remote_notes else get_release_notes(latest_tag or get_current_version())
            self.result.emit(False, latest_tag, notes, "")
        except Exception:
            self.result.emit(False, "", [], "")


class UpdateAvailableDialog(QDialog):
    """
    Modern modal dialog shown when an update is available.
    Displays version upgrade badge and comprehensive 2-3 line amendment bullets.
    """
    def __init__(self, current_ver: str, latest_ver: str, notes: List[str], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Update Available - PITC Suite")
        self.setFixedWidth(520)
        self.setModal(True)
        self.setWindowFlags(Qt.Dialog | Qt.WindowTitleHint | Qt.WindowCloseButtonHint)
        self.setStyleSheet("""
            QDialog {
                background: #1e222a;
                color: #edf0f5;
                font-family: 'Segoe UI', Tahoma, sans-serif;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 20, 22, 20)
        layout.setSpacing(14)

        # Header with icon and title
        header_layout = QHBoxLayout()
        header_layout.setSpacing(14)
        
        icon_lbl = QLabel("🚀")
        icon_lbl.setStyleSheet("""
            QLabel {
                font-size: 24pt;
                background: #2a313d;
                border: 1px solid #3b4252;
                border-radius: 26px;
                padding: 4px;
            }
        """)
        icon_lbl.setFixedSize(52, 52)
        icon_lbl.setAlignment(Qt.AlignCenter)
        header_layout.addWidget(icon_lbl)

        title_vbox = QVBoxLayout()
        title_vbox.setSpacing(4)
        title_lbl = QLabel("New Update Available")
        title_lbl.setStyleSheet("color: #60a5fa; font-size: 13pt; font-weight: bold;")
        title_vbox.addWidget(title_lbl)

        ver_badge = QLabel(f"Current: <span style='color: #94a3b8;'>{current_ver}</span> &nbsp;➔&nbsp; New: <b style='color: #4ade80;'>{latest_ver}</b>")
        ver_badge.setStyleSheet("color: #cbd5e1; font-size: 9.5pt;")
        title_vbox.addWidget(ver_badge)
        header_layout.addLayout(title_vbox)
        header_layout.addStretch()
        layout.addLayout(header_layout)

        # Separator line
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet("color: #334155; margin: 0;")
        layout.addWidget(sep)

        # What's New Section Box
        notes_box = QFrame()
        notes_box.setStyleSheet("""
            QFrame {
                background: #252b36;
                border: 1px solid #374151;
                border-radius: 8px;
            }
        """)
        notes_layout = QVBoxLayout(notes_box)
        notes_layout.setContentsMargins(14, 12, 14, 12)
        notes_layout.setSpacing(10)

        whats_new_title = QLabel("✨ What's New in this Update / New Amendments:")
        whats_new_title.setStyleSheet("color: #93c5fd; font-size: 9.5pt; font-weight: bold; border: none; background: transparent;")
        notes_layout.addWidget(whats_new_title)

        if not notes:
            notes = get_release_notes(latest_ver)

        for item in notes:
            bullet_row = QHBoxLayout()
            bullet_row.setSpacing(8)
            bullet_row.setAlignment(Qt.AlignTop)

            bullet_dot = QLabel("•")
            bullet_dot.setStyleSheet("color: #38bdf8; font-size: 11pt; font-weight: bold; border: none; background: transparent;")
            bullet_dot.setFixedWidth(12)
            bullet_row.addWidget(bullet_dot)

            bullet_text = QLabel(item)
            bullet_text.setWordWrap(True)
            bullet_text.setStyleSheet("color: #e2e8f0; font-size: 8.5pt; line-height: 1.4; border: none; background: transparent;")
            bullet_row.addWidget(bullet_text, 1)

            notes_layout.addLayout(bullet_row)

        layout.addWidget(notes_box)

        # Question prompt
        prompt_lbl = QLabel("Would you like to download and install this update now?")
        prompt_lbl.setStyleSheet("color: #94a3b8; font-size: 8.5pt;")
        layout.addWidget(prompt_lbl)

        # Action Buttons
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)
        btn_layout.addStretch()

        later_btn = QPushButton("Remind Me Later")
        later_btn.setCursor(Qt.PointingHandCursor)
        later_btn.setStyleSheet("""
            QPushButton {
                background: #334155;
                color: #cbd5e1;
                font-size: 9pt;
                padding: 8px 16px;
                border-radius: 6px;
                border: none;
            }
            QPushButton:hover {
                background: #475569;
                color: #ffffff;
            }
        """)
        later_btn.clicked.connect(self.reject)
        btn_layout.addWidget(later_btn)

        update_btn = QPushButton("Download & Install Now")
        update_btn.setCursor(Qt.PointingHandCursor)
        update_btn.setStyleSheet("""
            QPushButton {
                background: #16a34a;
                color: #ffffff;
                font-size: 9pt;
                font-weight: bold;
                padding: 8px 18px;
                border-radius: 6px;
                border: none;
            }
            QPushButton:hover {
                background: #15803d;
            }
        """)
        update_btn.clicked.connect(self.accept)
        btn_layout.addWidget(update_btn)

        layout.addLayout(btn_layout)


class UpToDateDialog(QDialog):
    """
    Dialog displayed when the app is up to date, or on post-update first launch,
    showcasing the current version's amendment notes.
    """
    def __init__(self, current_ver: str, notes: List[str], parent=None, title="PITC Suite - Up to Date", is_welcome=False):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setFixedWidth(500)
        self.setModal(True)
        self.setWindowFlags(Qt.Dialog | Qt.WindowTitleHint | Qt.WindowCloseButtonHint)
        self.setStyleSheet("""
            QDialog {
                background: #1e222a;
                color: #edf0f5;
                font-family: 'Segoe UI', Tahoma, sans-serif;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 20, 22, 20)
        layout.setSpacing(14)

        # Header
        header_layout = QHBoxLayout()
        header_layout.setSpacing(14)
        
        icon_char = "🎉" if is_welcome else "✅"
        icon_bg = "#1f3a2f" if is_welcome else "#1c352d"
        icon_lbl = QLabel(icon_char)
        icon_lbl.setStyleSheet(f"""
            QLabel {{
                font-size: 24pt;
                background: {icon_bg};
                border: 1px solid #334155;
                border-radius: 26px;
                padding: 4px;
            }}
        """)
        icon_lbl.setFixedSize(52, 52)
        icon_lbl.setAlignment(Qt.AlignCenter)
        header_layout.addWidget(icon_lbl)

        title_vbox = QVBoxLayout()
        title_vbox.setSpacing(4)
        header_title = f"Successfully Updated to {current_ver}" if is_welcome else "You are Running the Latest Version"
        title_lbl = QLabel(header_title)
        title_lbl.setStyleSheet("color: #4ade80; font-size: 12pt; font-weight: bold;")
        title_vbox.addWidget(title_lbl)

        sub_msg = "Here are the new amendments and improvements in this release:" if is_welcome else f"Version <b>{current_ver}</b> is the current release."
        ver_lbl = QLabel(sub_msg)
        ver_lbl.setStyleSheet("color: #94a3b8; font-size: 9pt;")
        title_vbox.addWidget(ver_lbl)
        header_layout.addLayout(title_vbox)
        header_layout.addStretch()
        layout.addLayout(header_layout)

        # Separator line
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet("color: #334155; margin: 0;")
        layout.addWidget(sep)

        # What's New in this Version box
        notes_box = QFrame()
        notes_box.setStyleSheet("""
            QFrame {
                background: #252b36;
                border: 1px solid #374151;
                border-radius: 8px;
            }
        """)
        notes_layout = QVBoxLayout(notes_box)
        notes_layout.setContentsMargins(14, 12, 14, 12)
        notes_layout.setSpacing(10)

        section_title = f"✨ What's New in {current_ver}:"
        whats_new_title = QLabel(section_title)
        whats_new_title.setStyleSheet("color: #93c5fd; font-size: 9.5pt; font-weight: bold; border: none; background: transparent;")
        notes_layout.addWidget(whats_new_title)

        if not notes:
            notes = get_release_notes(current_ver)

        for item in notes:
            bullet_row = QHBoxLayout()
            bullet_row.setSpacing(8)
            bullet_row.setAlignment(Qt.AlignTop)

            bullet_dot = QLabel("•")
            bullet_dot.setStyleSheet("color: #38bdf8; font-size: 11pt; font-weight: bold; border: none; background: transparent;")
            bullet_dot.setFixedWidth(12)
            bullet_row.addWidget(bullet_dot)

            bullet_text = QLabel(item)
            bullet_text.setWordWrap(True)
            bullet_text.setStyleSheet("color: #e2e8f0; font-size: 8.5pt; line-height: 1.4; border: none; background: transparent;")
            bullet_row.addWidget(bullet_text, 1)

            notes_layout.addLayout(bullet_row)

        layout.addWidget(notes_box)

        # OK button
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        ok_btn = QPushButton("Got it" if is_welcome else "OK")
        ok_btn.setCursor(Qt.PointingHandCursor)
        ok_btn.setStyleSheet("""
            QPushButton {
                background: #334155;
                color: #e2e8f0;
                font-size: 9pt;
                padding: 7px 22px;
                border-radius: 6px;
                border: none;
            }
            QPushButton:hover {
                background: #475569;
                color: #ffffff;
            }
        """)
        ok_btn.clicked.connect(self.accept)
        btn_layout.addWidget(ok_btn)
        layout.addLayout(btn_layout)


def show_post_update_dialog(parent_window, version_tag: str):
    """Shows the 'What's New' welcome dialog on first run after an update."""
    notes = get_release_notes(version_tag)
    dlg = UpToDateDialog(version_tag, notes, parent_window, title="What's New in this Version", is_welcome=True)
    dlg.exec()


def check_for_updates(parent_window, silent=False):
    """
    Asynchronously checks for application updates.
    In silent mode: notifies parent window if update is available.
    In interactive mode: shows UpdateAvailableDialog with What's New notes or UpToDateDialog.
    """
    checker = UpdateChecker(parent_window)
    parent_window._update_checker = checker  # Prevent GC
    
    def on_result(has_update, latest_tag, notes, download_url=""):
        if has_update:
            if silent:
                if hasattr(parent_window, 'notify_update_available'):
                    parent_window.notify_update_available(latest_tag, notes)
            else:
                dlg = UpdateAvailableDialog(get_current_version(), latest_tag, notes, parent_window)
                if dlg.exec() == QDialog.Accepted:
                    start_download(parent_window, latest_tag, notes, download_url=download_url)
        else:
            if not silent:
                if latest_tag == "":
                    QMessageBox.warning(parent_window, "Update Failed", "Could not check for updates. Please verify your internet connection.")
                else:
                    dlg = UpToDateDialog(get_current_version(), notes, parent_window)
                    dlg.exec()
    
    checker.result.connect(on_result)
    checker.start()


def start_download(parent_window, latest_tag=None, notes=None, download_url=None):
    """Downloads update setup installer and executes detached silent install."""
    progress_dialog = QProgressDialog("Downloading Update...", "Cancel", 0, 100, parent_window)
    ver_str = f" ({latest_tag})" if latest_tag else ""
    progress_dialog.setWindowTitle(f"Updating PITC Suite{ver_str}")
    progress_dialog.setWindowModality(Qt.WindowModal)
    progress_dialog.setAutoClose(True)
    progress_dialog.setAutoReset(True)
    progress_dialog.resize(420, 110)

    dest_path = os.path.join(os.environ.get("TEMP", "C:\\Temp"), "PITCSuite_Setup.exe")
    target_source = download_url if download_url else GDRIVE_SETUP_FILE_ID
    
    downloader = DownloadThread(target_source, dest_path)
    parent_window._downloader = downloader  # Prevent GC crash
    
    def on_progress(downloaded, total):
        if total > 0:
            percent = int((downloaded / total) * 100)
            progress_dialog.setValue(percent)
        else:
            progress_dialog.setValue(0)
            progress_dialog.setLabelText(f"Downloading... {downloaded // 1024 // 1024} MB")

    def on_finished(path):
        progress_dialog.close()
        
        msg = QMessageBox(parent_window)
        msg.setWindowTitle("Application Update Ready")
        msg.setIcon(QMessageBox.Information)
        msg.setText(f"The update{ver_str} is ready to install.")
        msg.setInformativeText(
            "The installer will now launch to install the update.\n\n"
            "All new amendments will be active immediately upon completion."
        )
        msg.addButton("Install Update", QMessageBox.AcceptRole)
        msg.exec()
        
        # Strip all PyInstaller, Python, and Tk environment variables to prevent child process conflicts
        env = os.environ.copy()
        keys_to_remove = [
            k for k in env
            if k.startswith('_PYI')
            or k.startswith('_MEI')
            or k.startswith('PYTHON')
            or k.startswith('TCL')
            or k.startswith('TK')
        ]
        for k in keys_to_remove:
            env.pop(k, None)
            
        if hasattr(sys, '_MEIPASS'):
            meipass = sys._MEIPASS.lower()
            path_dirs = env.get('PATH', '').split(os.pathsep)
            path_dirs = [p for p in path_dirs if p.lower().rstrip(r'\/') != meipass and '_mei' not in p.lower()]
            env['PATH'] = os.pathsep.join(path_dirs)
            
        # Launch Setup.exe with /SILENT /SP- /FORCECLOSEAPPLICATIONS /AUTOUPDATE=1
        # - /SILENT keeps the installation progress bar visible on screen while suppressing "Next" click wizard prompts
        # - /SP- disables the initial "Do you wish to continue?" prompt
        # - /FORCECLOSEAPPLICATIONS forcefully closes running applications using files that need updating
        # - /AUTOUPDATE=1 instructs installer to auto-relaunch pitcsuite.exe upon completion
        installer_args = [
            path,
            "/SILENT",
            "/SP-",
            "/FORCECLOSEAPPLICATIONS",
            "/AUTOUPDATE=1"
        ]

        creation_flags = 0
        if sys.platform == 'win32':
            creation_flags = subprocess.CREATE_NEW_PROCESS_GROUP
            if hasattr(subprocess, 'DETACHED_PROCESS'):
                creation_flags |= subprocess.DETACHED_PROCESS

        # Reliably spawn detached installer process directly using OS CreateProcessW
        try:
            subprocess.Popen(
                installer_args,
                env=env,
                creationflags=creation_flags,
                close_fds=True
            )
        except Exception:
            # Fallback to ShellExecute if CreateProcess direct fails
            if sys.platform == 'win32' and hasattr(os, 'startfile'):
                try:
                    os.startfile(path, "open", arguments=" ".join(installer_args[1:]))
                except Exception:
                    pass

        # Brief pause to ensure the OS process creation handshake completes fully
        time.sleep(0.4)

        # Forcefully exit process to release executable and DLL file locks instantly,
        # preventing "file in use" errors during update
        try:
            app_inst = QApplication.instance()
            if app_inst:
                app_inst.quit()
        except Exception:
            pass
        os._exit(0)

    def on_error(err):
        progress_dialog.close()
        QMessageBox.critical(parent_window, "Download Failed", f"An error occurred while downloading the update:\n{err}")

    downloader.progress.connect(on_progress)
    downloader.finished.connect(on_finished)
    downloader.error.connect(on_error)
    
    progress_dialog.canceled.connect(downloader.terminate)
    
    downloader.start()
    progress_dialog.show()

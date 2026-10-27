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
import tempfile
from typing import List, Dict, Tuple, Optional
from packaging import version
from PySide6.QtWidgets import (
    QMessageBox, QProgressDialog, QApplication, QDialog,
    QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QFrame,
    QScrollArea, QWidget, QProgressBar
)
from PySide6.QtCore import QThread, Signal, Qt, QObject, QTimer, QSettings
from PySide6.QtGui import QFont, QCursor

GDRIVE_VERSION_FILE_ID = "1gmIKNgsj2hdYrrQtiwXHyQcXzCDR-ZZb" 
GDRIVE_SETUP_FILE_ID = "1wghUxDo3ezed5szbpXeWtWlyof4ykBIu"

# Comprehensive 2-3 line release notes detailing recent amendments
DEFAULT_RELEASE_NOTES: Dict[str, List[str]] = {
    "v2.8.101": [
        "<b>PITC Downloader — Windows 11 Chrome Driver Stability:</b> Resolved browser restart failure after 35 AC numbers on Windows 11 by adding forceful process tree cleanup, lingering lock file clearing, and anti-background Chromium flags.",
        "<b>Robust Session Recovery & Auto-Save:</b> Added session-aware login detection and automated 3-attempt restart retry loop to prevent thread crashes, saving Excel progress before restarts.",
        "<b>Windows Efficiency Mode (EcoQoS) Handling:</b> Hardened Chrome shutdown and reopen pipeline against Windows 11 process throttling and file handle delays."
    ],
    "v2.8.100": [
        "<b>Bill Scrapper — Accuracy & Layout Parity:</b> Aligned bill fetching procedure with standalone Bill Scrapper architecture, fixing column offset on kWh present readings, multiplying factors (MF), and kWh units.",
        "<b>Urdu Heading Text Filtering:</b> Fixed extraction of status, name_address, and due_date to properly discard bilingual Urdu translation headings and capture actual bill values.",
        "<b>Total Units Precision:</b> Resolved Total Units summing multiplying factors instead of consumed units, ensuring accurate total consumption calculation."
    ],
    "v2.8.99": [
        "<b>E/Bills Downloader Crash Fix:</b> Resolved crash when processing bills caused by PyMuPDF 2.0 ONNX layout model lookup (missing layout_rf2.4.1+imf1.yaml resource).",
        "<b>Dependency & Build Stabilization:</b> Pinned PyMuPDF to <2.0.0 and actively excluded unused ONNX layout models from PyInstaller builds to keep binary lightweight and fast.",
        "<b>Resilient PDF Formatting:</b> Isolated A4 scaling and Sr. No. stamping to prevent individual PDF formatting issues from halting multi-account batch downloads."
    ],
    "v2.8.98": [
        "<b>E/Bills Downloader URL Persistence:</b> Bills URL edits now persist across app relaunches with default set to https://bill.pitc.com.pk/pescobill.",
        "<b>PITC Report Downloader DISCO Selection:</b> Added DISCO dropdown selector with PESCO as default; persists user selection across relaunches and automates DISCO selection & credentials login on cp.pitc.com.pk.",
        "<b>Bill Scrapper Custom Bills URL:</b> Added editable Bills URL option with system persistence for subsequent relaunches (defaults to https://bill.pitc.com.pk/pescobill)."
    ],
    "v2.8.97": [
        "<b>Ultra-Smooth Auto-Updater (Bill Scrapper Architecture):</b> Aligned installer and updater pipeline with Bill Scrapper — unique timestamped download paths, elimination of Restart Manager stalls, and instant process termination.",
        "<b>Zero Delays & Seamless Relaunch:</b> Removed blocking dialogs and timeouts on update completion, launching the installer detached and restarting into the latest release seamlessly.",
        "<b>Bill Scrapper — Billing History & Analysis:</b> Full 1-Year/6-Months history extraction, A-1A(01) Protected status calculation, and native numeric Excel formatting."
    ],
    "v2.8.96": [
        "<b>Auto-Updater Resilience & Direct Proxy Delivery:</b> Fixed update download URL forwarding to ensure silent and manual updates directly download the latest installer from the secure API proxy.",
        "<b>Dynamic API Fallback:</b> Added automated recovery to query the latest release endpoints directly if download URLs are omitted, permanently preventing stale installer update loops.",
        "<b>Bill Scrapper — Billing History & Details:</b> Retains full 1-Year/6-Months history extraction, A-1A(01) Protected status calculation, Total Units & Payment, and metadata fields."
    ],
    "v2.8.95": [
        "<b>Bill Scrapper — Billing/Payment History:</b> Added 1-Year (11 historical months) and 6-Months (5 historical months) history extraction options with Month, Status, Units, Bill, and Payment details.",
        "<b>Protected / Non-Protected Calculation:</b> Added automated Protected/Non-Protected status evaluation strictly for A-1A(01) tariff consumers (≤ 200 units across last 6 months), keeping all other tariffs blank.",
        "<b>Total Units & Total Payment:</b> Added Total Payment and Total Units columns with native numeric formatting placed right next to each other.",
        "<b>BNo, SDiv & Meter Type:</b> Added BNo (batch no), SDiv (5-digit sub-division code), and Meter Type (S-P / 3-P) extraction alongside AutoFilter support on Excel headings."
    ],
    "v2.8.94": [
        "<b>Security Hardening (Nuitka Native C++ Compilation):</b> Proprietary tools, licensing, and business logic compiled into native C++ machine code binaries (.pyd) to protect against decompilation and reverse-engineering.",
        "<b>CP-139 Posting Checker/PDF:</b> Added PDF Generation (single & separate PDFs named by Para Sr. No.), automatic Column C Remarks updating ('Found and PDF made' / 'AC No Not found'), and template download.",
        "<b>Bill Scrapper:</b> Added FPA and Current Bill fetching for comprehensive consumer billing data extraction."
    ],
    "v2.8.93": [
        "<b>Security Hardening (Nuitka Native C++ Compilation):</b> Proprietary tools, licensing, and business logic compiled into native C++ machine code binaries (.pyd) to protect against decompilation and reverse-engineering.",
        "<b>CP-139 Posting Checker/PDF:</b> Added PDF Generation (single & separate PDFs named by Para Sr. No.), automatic Column C Remarks updating ('Found and PDF made' / 'AC No Not found'), and template download.",
        "<b>Bill Scrapper:</b> Added FPA and Current Bill fetching for comprehensive consumer billing data extraction."
    ],
    "v2.8.92": [
        "<b>CP-139 Posting Checker/PDF:</b> Added PDF Generation",
        "<b>Bill Scrapper:</b> Added FPA and Current Bill fetching"
    ],
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
        return "v2.8.93"


def get_release_notes(ver: str) -> List[str]:
    """Retrieve comprehensive release notes bullets for a given version tag."""
    clean_ver = ver.strip()
    if clean_ver in DEFAULT_RELEASE_NOTES and DEFAULT_RELEASE_NOTES[clean_ver]:
        return DEFAULT_RELEASE_NOTES[clean_ver]
    with_v = f"v{clean_ver.lstrip('vV')}"
    if with_v in DEFAULT_RELEASE_NOTES and DEFAULT_RELEASE_NOTES[with_v]:
        return DEFAULT_RELEASE_NOTES[with_v]
    
    # Fallback to the latest available notes (keeps previous what's new if version unchanged)
    for k in sorted(DEFAULT_RELEASE_NOTES.keys(), reverse=True):
        if DEFAULT_RELEASE_NOTES[k]:
            return DEFAULT_RELEASE_NOTES[k]
    return []


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

    def __init__(self, target: str, dest: str, expected_size: int = 0):
        super().__init__()
        self.target = target
        self.dest = dest
        self.expected_size = expected_size
        self._is_cancelled = False

    def cancel(self):
        self._is_cancelled = True

    def run(self):
        try:
            # 1. Direct or Vercel Proxy streaming URL (HTTP/HTTPS)
            if self.target and (self.target.startswith("http://") or self.target.startswith("https://")) and "docs.google.com" not in self.target and "drive.google.com" not in self.target:
                session = requests.Session()
                headers = {
                    "User-Agent": "PITCSuite-DesktopApp",
                    "Accept-Encoding": "identity",
                }
                response = session.get(self.target, headers=headers, stream=True, timeout=30)
                response.raise_for_status()

                # Extract total_size reliably
                total_size = 0
                cl = response.headers.get('content-length')
                if cl:
                    try:
                        total_size = int(cl)
                    except (ValueError, TypeError):
                        pass
                if total_size <= 0:
                    cr = response.headers.get('content-range')
                    if cr and '/' in cr:
                        try:
                            total_size = int(cr.split('/')[-1].strip())
                        except (ValueError, TypeError):
                            pass
                if total_size <= 0 and self.expected_size > 0:
                    total_size = int(self.expected_size)

                # Fallback to quick HEAD request with identity encoding
                if total_size <= 0:
                    try:
                        head_resp = session.head(self.target, headers=headers, allow_redirects=True, timeout=6)
                        head_cl = head_resp.headers.get('content-length')
                        if head_cl:
                            total_size = int(head_cl)
                    except Exception:
                        pass

                # If total_size is still unknown, fallback to typical installer package size (~130.5 MB)
                if total_size <= 0:
                    total_size = 136908774

                downloaded = 0
                self.progress.emit(0, total_size)
                with open(self.dest, "wb") as f:
                    for chunk in response.iter_content(chunk_size=65536):
                        if self._is_cancelled:
                            break
                        if chunk:
                            f.write(chunk)
                            downloaded += len(chunk)
                            self.progress.emit(downloaded, total_size)

                if self._is_cancelled:
                    try:
                        if os.path.exists(self.dest):
                            os.remove(self.dest)
                    except Exception:
                        pass
                    return

                self.finished.emit(self.dest)
                return

            # 2. Google Drive File ID / URL fallback
            file_id = self.target if self.target and not self.target.startswith("http") else GDRIVE_SETUP_FILE_ID
            url = "https://docs.google.com/uc?export=download"
            session = requests.Session()
            headers = {
                "User-Agent": "PITCSuite-DesktopApp",
                "Accept-Encoding": "identity",
            }
            response = session.get(url, params={'id': file_id}, headers=headers, stream=True)
            
            if "text/html" in response.headers.get("Content-Type", ""):
                text = response.text
                match = re.search(r'name="uuid" value="([^"]+)"', text)
                if match:
                    uuid_val = match.group(1)
                    final_url = f"https://drive.usercontent.google.com/download?id={file_id}&export=download&confirm=t&uuid={uuid_val}"
                    response = session.get(final_url, headers=headers, stream=True)
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
                    response = session.get(final_url, headers=headers, stream=True)

            response.raise_for_status()

            total_size = 0
            cl = response.headers.get('content-length')
            if cl:
                try:
                    total_size = int(cl)
                except (ValueError, TypeError):
                    pass
            if total_size <= 0:
                cr = response.headers.get('content-range')
                if cr and '/' in cr:
                    try:
                        total_size = int(cr.split('/')[-1].strip())
                    except (ValueError, TypeError):
                        pass
            if total_size <= 0 and self.expected_size > 0:
                total_size = int(self.expected_size)
            if total_size <= 0:
                total_size = 136908774

            downloaded = 0
            self.progress.emit(0, total_size)
            with open(self.dest, "wb") as file:
                for chunk in response.iter_content(chunk_size=16384):
                    if self._is_cancelled:
                        break
                    if chunk:
                        file.write(chunk)
                        downloaded += len(chunk)
                        self.progress.emit(downloaded, total_size)

            if self._is_cancelled:
                try:
                    if os.path.exists(self.dest):
                        os.remove(self.dest)
                except Exception:
                    pass
                return

            self.finished.emit(self.dest)
        except Exception as e:
            self.error.emit(str(e))


class UpdateChecker(QThread):
    result = Signal(bool, str, list, str, int)  # update_available, version_string, release_notes, download_url, file_size

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
                    file_size = int(data.get("file_size", 0) or 0)
                    server_notes = data.get("release_notes", "")

                    notes_list = []
                    raw_notes = []
                    if isinstance(server_notes, list):
                        raw_notes = server_notes
                    elif isinstance(server_notes, str) and server_notes.strip():
                        raw_notes = server_notes.splitlines()

                    for n in raw_notes:
                        line_str = str(n).strip()
                        # Strictly reject GitHub compare links, changelog links, or commit diff urls
                        if re.search(r'(?:full\s*changelog|github\.com/.+/compare|compare/v\d|\/compare\/)', line_str, re.I):
                            continue
                        line_clean = re.sub(r'^[•\-\*\d\.\)\s]+', '', line_str).strip()
                        if line_clean:
                            notes_list.append(line_clean)

                    # If no valid notes remain from the server, fallback to local/previous release notes
                    if not notes_list:
                        notes_list = get_release_notes(clean_tag)

                    v_latest = version.parse(clean_tag.lstrip('vV'))
                    v_current = version.parse(get_current_version().lstrip('vV'))
                    if v_latest > v_current:
                        self.result.emit(True, clean_tag, notes_list, download_url, file_size)
                        return
                    else:
                        self.result.emit(False, clean_tag, notes_list, download_url, file_size)
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
                clean_remote = [
                    n for n in remote_notes 
                    if not re.search(r'(?:full\s*changelog|github\.com/.+/compare|compare/v\d|\/compare\/)', n, re.I)
                ] if remote_notes else []
                notes = clean_remote if clean_remote else get_release_notes(latest_tag)
                if v_latest > v_current:
                    self.result.emit(True, latest_tag, notes, GDRIVE_SETUP_FILE_ID, 0)
                    return
                else:
                    self.result.emit(False, latest_tag, notes, "", 0)
                    return
            clean_remote = [
                n for n in remote_notes 
                if not re.search(r'(?:full\s*changelog|github\.com/.+/compare|compare/v\d|\/compare\/)', n, re.I)
            ] if remote_notes else []
            notes = clean_remote if clean_remote else get_release_notes(latest_tag or get_current_version())
            self.result.emit(False, latest_tag, notes, "", 0)
        except Exception:
            self.result.emit(False, "", [], "", 0)


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
    
    def on_result(has_update, latest_tag, notes, download_url="", file_size=0):
        if has_update:
            if silent:
                if hasattr(parent_window, 'notify_update_available'):
                    try:
                        parent_window.notify_update_available(latest_tag, notes, download_url=download_url, file_size=file_size)
                    except TypeError:
                        try:
                            parent_window.notify_update_available(latest_tag, notes, download_url=download_url)
                        except TypeError:
                            parent_window.notify_update_available(latest_tag, notes)
            else:
                dlg = UpdateAvailableDialog(get_current_version(), latest_tag, notes, parent_window)
                if dlg.exec() == QDialog.Accepted:
                    start_download(parent_window, latest_tag, notes, download_url=download_url, file_size=file_size)
        else:
            if not silent:
                if latest_tag == "":
                    QMessageBox.warning(parent_window, "Update Failed", "Could not check for updates. Please verify your internet connection.")
                else:
                    dlg = UpToDateDialog(get_current_version(), notes, parent_window)
                    dlg.exec()
    
    checker.result.connect(on_result)
    checker.start()


class UpdateDownloadDialog(QDialog):
    """
    Dedicated modal dialog displaying live download progress and applying the update.
    Matches Bill Scrapper's UpdateDialog architecture:
    - Vibrant green progress bar with smooth percentage text
    - Dual status labels: activity ('Downloading update package...') and statistics ('45% (61.5 MB / 136.9 MB)')
    - Automatic fallback size estimation so progress bar NEVER gets stuck at 0%
    - Detached silent installation with multi-tier process termination and seamless relaunch
    """
    def __init__(self, version_tag: str = "", file_size: int = 0, parent=None):
        super().__init__(parent)
        self.version_tag = version_tag
        self.total_size = int(file_size or 0)
        self.is_cancelled = False
        self.downloader: Optional[DownloadThread] = None

        ver_str = f" ({version_tag})" if version_tag else ""
        self.setWindowTitle(f"Updating PITC Suite{ver_str}")
        self.setFixedWidth(460)
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
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(10)

        # Header Title
        title_text = f"Downloading Update{ver_str}..."
        self.lbl_title = QLabel(title_text)
        self.lbl_title.setStyleSheet("color: #60a5fa; font-size: 11pt; font-weight: bold;")
        layout.addWidget(self.lbl_title)

        # Status Label
        self.lbl_status = QLabel("Connecting to update server...")
        self.lbl_status.setStyleSheet("color: #cbd5e1; font-size: 9pt;")
        layout.addWidget(self.lbl_status)

        # Styled Green Progress Bar (matching Bill Scrapper)
        self.bar = QProgressBar()
        self.bar.setRange(0, 100)
        self.bar.setValue(0)
        self.bar.setFixedHeight(20)
        self.bar.setTextVisible(True)
        self.bar.setStyleSheet("""
            QProgressBar {
                background-color: #161a21;
                border: 1px solid #374151;
                border-radius: 5px;
                text-align: center;
                color: #ffffff;
                font-size: 8.5pt;
                font-weight: bold;
            }
            QProgressBar::chunk {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #16a34a, stop:1 #22c55e);
                border-radius: 4px;
            }
        """)
        layout.addWidget(self.bar)

        # Details label
        init_mb = (self.total_size / (1024 * 1024)) if self.total_size > 0 else 0.0
        details_init = f"0% (0.0 MB / {init_mb:.1f} MB)" if init_mb > 0 else "0% (0.0 MB)"
        self.lbl_details = QLabel(details_init)
        self.lbl_details.setStyleSheet("color: #94a3b8; font-size: 8.5pt;")
        layout.addWidget(self.lbl_details)

        # Bottom Button Row
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.setCursor(Qt.PointingHandCursor)
        self.cancel_btn.setStyleSheet("""
            QPushButton {
                background: #334155;
                color: #cbd5e1;
                font-size: 9pt;
                padding: 6px 16px;
                border-radius: 5px;
                border: none;
            }
            QPushButton:hover {
                background: #475569;
                color: #ffffff;
            }
            QPushButton:disabled {
                background: #252b36;
                color: #64748b;
            }
        """)
        self.cancel_btn.clicked.connect(self.on_cancel)
        btn_layout.addWidget(self.cancel_btn)
        layout.addLayout(btn_layout)

    def on_progress(self, downloaded: int, total: int):
        if self.is_cancelled:
            return

        if total > 0:
            self.total_size = total
        elif self.total_size <= 0:
            self.total_size = 136908774

        dl_mb = downloaded / (1024 * 1024)
        tot_mb = self.total_size / (1024 * 1024)
        pct = min(100, max(0, int(round((downloaded / self.total_size) * 100))))

        self.bar.setValue(pct)
        self.lbl_status.setText("Downloading update package...")
        self.lbl_details.setText(f"{pct}% ({dl_mb:.1f} MB / {tot_mb:.1f} MB)")

    def on_finished(self, path: str):
        self.bar.setValue(100)
        self.lbl_status.setText("Download complete! Restarting to apply update...")
        self.lbl_details.setText("Please wait a moment...")
        self.cancel_btn.setEnabled(False)
        QApplication.processEvents()

        # Smooth automatic restart with 600ms delay, exactly matching BillScrapper
        QTimer.singleShot(600, lambda: self._apply_and_restart(path))

    def on_error(self, err_msg: str):
        self.close()
        QMessageBox.critical(self.parent(), "Download Failed", f"An error occurred while downloading the update:\n{err_msg}")

    def on_cancel(self):
        self.is_cancelled = True
        self.lbl_status.setText("Cancelling download...")
        self.cancel_btn.setEnabled(False)
        if self.downloader:
            self.downloader.cancel()
            self.downloader.terminate()
        self.close()

    def closeEvent(self, event):
        if not self.is_cancelled:
            self.on_cancel()
        event.accept()

    def _apply_and_restart(self, path: str):
        # Strip all PyInstaller, Python, and Tk environment variables to prevent child process conflicts
        clean_env = os.environ.copy()
        keys_to_remove = [
            k for k in clean_env
            if k.startswith('_PYI')
            or k.startswith('_MEI')
            or k.startswith('PYTHON')
            or k.startswith('TCL')
            or k.startswith('TK')
        ]
        for k in keys_to_remove:
            clean_env.pop(k, None)

        if hasattr(sys, '_MEIPASS'):
            meipass = sys._MEIPASS.lower()
            path_dirs = clean_env.get('PATH', '').split(os.pathsep)
            path_dirs = [p for p in path_dirs if p.lower().rstrip(r'\/') != meipass and '_mei' not in p.lower()]
            clean_env['PATH'] = os.pathsep.join(path_dirs)

        # Launch Setup.exe matching BillScrapper:
        # /SILENT /CLOSEAPPLICATIONS /FORCECLOSEAPPLICATIONS /AUTOUPDATE=1
        installer_args = [
            path,
            "/SILENT",
            "/CLOSEAPPLICATIONS",
            "/FORCECLOSEAPPLICATIONS",
            "/AUTOUPDATE=1"
        ]

        # Use exact detached creation flags from BillScrapper (CREATE_NO_WINDOW | CREATE_NEW_PROCESS_GROUP)
        DETACHED_FLAGS = 0x08000000 | 0x00000200
        try:
            subprocess.Popen(
                installer_args,
                env=clean_env,
                creationflags=DETACHED_FLAGS,
                close_fds=True
            )
        except Exception:
            subprocess.Popen(
                installer_args,
                env=clean_env,
                close_fds=True
            )

        # Immediately close window and exit process cleanly without hanging on Qt event loop or threads
        parent = self.parent()
        if parent:
            try:
                parent.close()
            except Exception:
                pass
        os._exit(0)


def start_download(parent_window, latest_tag=None, notes=None, download_url=None, file_size=0):
    """Downloads update setup installer and executes detached silent install matching BillScrapper."""
    # Unique timestamped installer path to prevent file locking and cache collisions (matching BillScrapper)
    fname = f"update_{int(time.time())}_PITCSuite_Setup.exe"
    dest_path = os.path.join(tempfile.gettempdir(), fname)

    # Resolve target_source: prefer direct/API download_url over obsolete Google Drive fallback
    target_source = download_url
    resolved_size = file_size
    if not target_source or target_source == GDRIVE_SETUP_FILE_ID:
        try:
            from pitcsuite.license_manager import get_api_base_url
            api_url = get_api_base_url()
            resp = requests.get(f"{api_url}/api/updates/latest", timeout=6)
            if resp.status_code == 200:
                data = resp.json()
                remote_url = data.get("download_url")
                if remote_url and remote_url.startswith("http"):
                    target_source = remote_url
                if not resolved_size:
                    resolved_size = int(data.get("file_size", 0) or 0)
        except Exception:
            pass

    if not target_source:
        target_source = GDRIVE_SETUP_FILE_ID

    dialog = UpdateDownloadDialog(version_tag=latest_tag or "", file_size=resolved_size, parent=parent_window)
    downloader = DownloadThread(target_source, dest_path, expected_size=resolved_size)
    dialog.downloader = downloader
    parent_window._downloader = downloader  # Prevent GC crash
    parent_window._download_dialog = dialog

    downloader.progress.connect(dialog.on_progress)
    downloader.finished.connect(dialog.on_finished)
    downloader.error.connect(dialog.on_error)

    downloader.start()
    dialog.show()

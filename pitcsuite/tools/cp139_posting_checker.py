"""
CP 139 Posting Checker Engine & GUI for PITC Suite
Parses WAPDA RCP-139 Unidentified Cash / Debtor's Adjustment posting reports (.txt and .pdf),
extracts 14-digit Ref#, S/Div (digits 3-7), Grand Total amount, and posting status/errors (Posted / Not Pos),
supports deduplication (single postings) and exports professionally styled Excel (.xlsx) workbooks.
"""

import os
import re
import glob
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any, Tuple, Optional

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QLineEdit,
    QListWidget, QGroupBox, QComboBox, QCheckBox, QTextEdit,
    QFileDialog, QMessageBox, QFrame
)


from pitcsuite.ui_helpers import START_BTN, STOP_BTN, hline, lbl, notify


def _match_record_start(line_str: str) -> Optional[str]:
    """
    Check if a line represents the start of a record row.
    Returns the reference number (13 or 14 digits) or None.
    Handles:
      - Format 1 (Posting List): begins directly with Ref# (13 or 14 digits)
      - Format 2 (Proof List / Adjustor): begins with RecNo (1-6 digits) followed by Ref# (13 or 14 digits)
    """
    if any(h in line_str for h in (
        "Report", "PROCESSING", "DIVISION", "TRANSACTION", "TRANS.",
        "RecNo.", "Rec #", "DATE & TIME", "Statistical Summary", "TOTAL"
    )):
        return None

    # Format 2: RecNo (1-6 digits) followed by whitespace and 13-14 digit Ref#
    m2 = re.match(r"^\s*\d{1,6}\s+(\d{13,14})\b", line_str)
    if m2:
        return m2.group(1)

    # Format 1: Starts directly with 13-14 digit Ref#
    m1 = re.match(r"^\s*(\d{13,14})\b", line_str)
    if m1:
        return m1.group(1)

    return None


def parse_cp139_file(filepath: str) -> List[Dict[str, Any]]:
    """
    Parse a single CP-139 posting report text (.txt) or PDF (.pdf) file.
    Supports both Posting List format (with 14-digit Ref#) and Proof List / Adjustor format
    (with RecNo, 13/14-digit Ref#, RFP/NR status flags).
    Returns list of dicts with keys:
      - 'Ref#': 13 or 14-digit string
      - 'S/Div': 5-digit sub-division string
      - 'G. Total': numeric float/int
      - 'Errors': 'Posted' or 'Not Pos'
      - 'Error_Detail': status with error code if found, e.g. 'Not Pos (G09)' or 'Not Pos (NPV)'
      - 'Source_File': filename
    """
    records = []
    if not os.path.isfile(filepath):
        if os.path.isfile(filepath + ".txt"):
            filepath = filepath + ".txt"
        elif os.path.isfile(filepath + ".pdf"):
            filepath = filepath + ".pdf"
        else:
            return records

    filename = os.path.basename(filepath)
    ext = os.path.splitext(filepath)[1].lower()

    lines = []
    if ext == ".pdf":
        try:
            import pypdf
            reader = pypdf.PdfReader(filepath)
            for page in reader.pages:
                txt = page.extract_text() or ""
                lines.extend(txt.splitlines(keepends=True))
        except Exception:
            lines = []
    else:
        with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
            lines = f.readlines()

    n = len(lines)
    i = 0
    while i < n:
        line = lines[i]
        ref = _match_record_start(line)
        if ref:
            s_div = ref[2:7]  # WAPDA sub-division is digits 3 to 7

            # G. Total is in line i+1 (under Ins. Bill Amt / G. Total header)
            g_total_val = 0
            if i + 1 < n:
                l_next = lines[i + 1]
                # Slice columns around 60:85
                raw_gt = l_next[60:85].strip() if len(l_next) > 60 else ""
                m_gt = re.search(r"[-+]?[\d,]+(?:\.\d+)?", raw_gt)
                if not m_gt:
                    nums = re.findall(r"[-+]?[\d,]+(?:\.\d+)?", l_next.strip())
                    if len(nums) >= 2:
                        raw_gt = nums[1]
                        m_gt = re.search(r"[-+]?[\d,]+(?:\.\d+)?", raw_gt)
                    elif len(nums) == 1:
                        raw_gt = nums[0]
                        m_gt = re.search(r"[-+]?[\d,]+(?:\.\d+)?", raw_gt)

                if m_gt:
                    clean_str = m_gt.group(0).replace(",", "")
                    try:
                        f_val = float(clean_str)
                        g_total_val = int(f_val) if f_val.is_integer() else round(f_val, 2)
                    except ValueError:
                        g_total_val = clean_str

            # Errors: check lines from i to i+7 for 'Posted', 'RFP', 'Not Pos', or 'NR'
            # Stop if another record begins or page break occurs
            status = "Posted"  # default
            err_code = ""
            for j in range(i, min(i + 8, n)):
                if j > i:
                    if _match_record_start(lines[j]) or "\x0c" in lines[j] or "Report" in lines[j] or "Statistical Summary" in lines[j]:
                        break

                row_str = lines[j]
                if "Not Pos" in row_str or re.search(r"\bNR\b", row_str):
                    status = "Not Pos"
                    # Check next line for error code such as G09 (Format 1)
                    if "Not Pos" in row_str and j + 1 < n:
                        code_m = re.search(r"\b([A-Z]\d{2,3})\b", lines[j + 1])
                        if code_m:
                            err_code = code_m.group(1)
                    if not err_code:
                        # Check same line and next lines for error tokens (Format 2: e.g. R27,R23,NPV, NPV, DUP)
                        for k in range(j, min(j + 2, n)):
                            cand = lines[k]
                            if k == j:
                                cand = re.sub(r"\bNR\b", "", cand)
                            code_m = re.search(r"\b([A-Z][A-Z0-9]*(?:,[A-Z0-9]+)*)\b", cand)
                            if code_m and code_m.group(1) not in (
                                "REPORT", "RCP", "DATE", "DIVISION", "TRANSACTION", "TOTAL", "ED", "CD", "GD"
                            ):
                                err_code = code_m.group(1)
                                break
                    break
                elif "Posted" in row_str or re.search(r"\bRFP\b", row_str):
                    status = "Posted"

            detail = f"{status} ({err_code})" if err_code else status

            records.append({
                "Ref#": ref,
                "S/Div": s_div,
                "G. Total": g_total_val,
                "Errors": status,
                "Error_Detail": detail,
                "Source_File": filename,
            })
            i += 1
        else:
            i += 1

    return records


def deduplicate_records(records: List[Dict[str, Any]], mode: str = "prefer_posted") -> Tuple[List[Dict[str, Any]], int]:
    """
    Deduplicates records by 'Ref#'.
    Modes:
      - 'prefer_posted': If multiple records exist for a Ref#, keep 'Posted' if available; otherwise first.
      - 'first': Strictly keep the first occurrence encountered.
      - 'last': Keep the last occurrence encountered.
    Returns (deduped_records, duplicates_removed_count).
    """
    if mode == "none" or not mode:
        return records, 0

    seen = {}
    dups_count = 0

    if mode == "last":
        # Keep the latest occurrence
        for r in records:
            ref = r["Ref#"]
            if ref in seen:
                dups_count += 1
            seen[ref] = r
        return list(seen.values()), dups_count

    for r in records:
        ref = r["Ref#"]
        if ref not in seen:
            seen[ref] = r
        else:
            dups_count += 1
            if mode == "prefer_posted":
                existing = seen[ref]
                if existing.get("Errors") != "Posted" and r.get("Errors") == "Posted":
                    seen[ref] = r

    return list(seen.values()), dups_count


def filter_records(records: List[Dict[str, Any]], status_filter: str = "all") -> List[Dict[str, Any]]:
    """
    Filter records based on posting status: 'all', 'posted', or 'not_pos'.
    """
    if status_filter == "posted":
        return [r for r in records if r.get("Errors") == "Posted"]
    elif status_filter == "not_pos":
        return [r for r in records if "Not Pos" in str(r.get("Errors", ""))]
    return records


def export_cp139_to_excel(
    records: List[Dict[str, Any]],
    output_path: str,
    include_source_file: bool = False
) -> Dict[str, Any]:
    """
    Exports parsed CP-139 records to a clean, professionally formatted Excel workbook (.xlsx).
    Ref# and S/Div are preserved as text to retain leading zeros.
    G. Total is formatted as numeric currency.
    Errors column is styled with soft green for Posted and soft red for Not Pos.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CP139_Postings"

    # Styling elements
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Segoe UI", size=10)
    total_font = Font(name="Segoe UI", size=10, bold=True)
    posted_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")  # soft green
    not_pos_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid") # soft red
    total_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")

    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )
    total_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="475569"),
        bottom=Side(style="double", color="475569")
    )

    headers = ["Ref#", "S/Div", "G. Total", "Errors"]
    if include_source_file:
        headers.append("Source File")

    ws.append(headers)

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    ws.row_dimensions[1].height = 26

    posted_count = 0
    not_pos_count = 0
    total_amount = 0.0

    for r_idx, rec in enumerate(records, start=2):
        # Ref#
        c_ref = ws.cell(row=r_idx, column=1, value=str(rec["Ref#"]))
        c_ref.number_format = "@"
        c_ref.font = data_font
        c_ref.alignment = Alignment(horizontal="center", vertical="center")
        c_ref.border = thin_border

        # S/Div
        c_sdiv = ws.cell(row=r_idx, column=2, value=str(rec["S/Div"]))
        c_sdiv.number_format = "@"
        c_sdiv.font = data_font
        c_sdiv.alignment = Alignment(horizontal="center", vertical="center")
        c_sdiv.border = thin_border

        # G. Total
        gt_val = rec["G. Total"]
        c_gt = ws.cell(row=r_idx, column=3, value=gt_val)
        if isinstance(gt_val, (int, float)):
            total_amount += gt_val
            c_gt.number_format = "#,##0.00" if isinstance(gt_val, float) else "#,##0"
        c_gt.font = data_font
        c_gt.alignment = Alignment(horizontal="right", vertical="center")
        c_gt.border = thin_border

        # Errors
        err_val = rec.get("Error_Detail") or rec.get("Errors", "")
        c_err = ws.cell(row=r_idx, column=4, value=err_val)
        c_err.font = data_font
        c_err.alignment = Alignment(horizontal="center", vertical="center")
        c_err.border = thin_border
        if rec.get("Errors") == "Posted":
            c_err.fill = posted_fill
            posted_count += 1
        elif "Not Pos" in rec.get("Errors", ""):
            c_err.fill = not_pos_fill
            not_pos_count += 1

        # Source File
        if include_source_file:
            c_src = ws.cell(row=r_idx, column=5, value=rec.get("Source_File", ""))
            c_src.font = data_font
            c_src.alignment = Alignment(horizontal="left", vertical="center")
            c_src.border = thin_border

        ws.row_dimensions[r_idx].height = 20

    # Add Summary Total Row if there are records
    last_row = len(records) + 1
    if records:
        sum_row = last_row + 1
        c_lbl = ws.cell(row=sum_row, column=1, value="TOTAL")
        c_lbl.font = total_font
        c_lbl.fill = total_fill
        c_lbl.alignment = Alignment(horizontal="center", vertical="center")
        c_lbl.border = total_border

        c_sdiv_tot = ws.cell(row=sum_row, column=2, value=f"{len(records)} records")
        c_sdiv_tot.font = total_font
        c_sdiv_tot.fill = total_fill
        c_sdiv_tot.alignment = Alignment(horizontal="center", vertical="center")
        c_sdiv_tot.border = total_border

        c_amt = ws.cell(row=sum_row, column=3, value=f"=SUM(C2:C{last_row})")
        c_amt.font = total_font
        c_amt.fill = total_fill
        c_amt.number_format = "#,##0.00" if isinstance(total_amount, float) else "#,##0"
        c_amt.alignment = Alignment(horizontal="right", vertical="center")
        c_amt.border = total_border

        c_stat = ws.cell(row=sum_row, column=4, value=f"Posted: {posted_count} | Not: {not_pos_count}")
        c_stat.font = total_font
        c_stat.fill = total_fill
        c_stat.alignment = Alignment(horizontal="center", vertical="center")
        c_stat.border = total_border

        if include_source_file:
            c_end = ws.cell(row=sum_row, column=5, value="")
            c_end.fill = total_fill
            c_end.border = total_border

        ws.row_dimensions[sum_row].height = 22

    # Auto-adjust column widths
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or "")
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 5, 14)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb.save(output_path)

    return {
        "total": len(records),
        "posted": posted_count,
        "not_pos": not_pos_count,
        "total_amount": total_amount,
        "saved_path": output_path
    }


class CP139Worker(QThread):
    """Background worker for parsing CP-139 posting files without freezing the GUI."""
    log = Signal(str)
    progress = Signal(int, int, str)
    done = Signal(str, dict, list)
    failed = Signal(str)

    def __init__(
        self,
        file_paths: List[str],
        output_path: str,
        deduplicate: bool = True,
        dedup_mode: str = "prefer_posted",
        status_filter: str = "all",
        include_source_file: bool = False
    ):
        super().__init__()
        self.file_paths = file_paths
        self.output_path = output_path
        self.deduplicate = deduplicate
        self.dedup_mode = dedup_mode
        self.status_filter = status_filter
        self.include_source_file = include_source_file

    def run(self):
        try:
            total_files = len(self.file_paths)
            self.log.emit(f"🚀 Starting CP-139 extraction across {total_files} file(s)...")

            all_records = []
            for idx, fp in enumerate(self.file_paths, start=1):
                if self.isInterruptionRequested():
                    self.log.emit("⏹ Operation cancelled by user.")
                    return

                fname = os.path.basename(fp)
                self.progress.emit(idx, total_files, f"Parsing {fname} ({idx}/{total_files})...")
                self.log.emit(f"📄 Reading: {fname}")
                recs = parse_cp139_file(fp)
                self.log.emit(f"   ↳ Extracted {len(recs)} posting record(s) from {fname}")
                all_records.extend(recs)

            total_raw = len(all_records)
            self.log.emit(f"\n📊 Total raw records extracted: {total_raw}")

            # Deduplication
            dups_count = 0
            if self.deduplicate:
                self.log.emit(f"🔍 Applying deduplication (Mode: {self.dedup_mode})...")
                all_records, dups_count = deduplicate_records(all_records, mode=self.dedup_mode)
                self.log.emit(f"   ↳ Retained {len(all_records)} unique records ({dups_count} duplicates removed).")

            # Status filtering
            if self.status_filter != "all":
                self.log.emit(f"🔍 Applying status filter ({self.status_filter})...")
                all_records = filter_records(all_records, self.status_filter)
                self.log.emit(f"   ↳ Retained {len(all_records)} records matching filter.")

            if not all_records:
                raise ValueError("No records found to export from the selected file(s).")

            # Export to Excel
            self.log.emit(f"\n💾 Generating structured Excel report: {self.output_path}...")
            stats = export_cp139_to_excel(
                all_records,
                self.output_path,
                include_source_file=self.include_source_file
            )
            stats["duplicates_removed"] = dups_count
            stats["raw_total"] = total_raw

            self.log.emit(f"✔ Export completed successfully!")
            self.log.emit(f"   • Total records in Excel: {stats['total']}")
            self.log.emit(f"   • Posted: {stats['posted']}")
            self.log.emit(f"   • Not Pos: {stats['not_pos']}")
            self.log.emit(f"   • Grand Total Sum: Rs. {stats['total_amount']:,.2f}")
            self.log.emit(f"   • Output file: {self.output_path}")

            self.done.emit(self.output_path, stats, all_records)
        except Exception as exc:
            self.failed.emit(str(exc))


class CP139PostingCheckerPanel(QWidget):
    """UI Panel for CP-139 Posting Checker in PITC Suite."""
    def __init__(self):
        super().__init__()
        self.worker = None
        self.file_paths = []
        self.last_output = ""

        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        # Header Title and Description
        layout.addWidget(lbl("CP 139 Posting Checker", bold=True, color="#7eb8f7"))
        layout.addWidget(lbl(
            "Extract Ref#, S/Div, Grand Total, and Posting Status from WAPDA RCP-139 reports into structured Excel.",
            color="#556"
        ))
        layout.addWidget(hline())

        # Inputs Group
        group_in = QGroupBox("Source Reports (WAPDA RCP-139 .txt / .pdf files)")
        form_in = QVBoxLayout(group_in)
        self.file_list = QListWidget()
        self.file_list.setMaximumHeight(75)
        self.file_list.setMinimumHeight(55)
        form_in.addWidget(self.file_list)

        btn_row = QHBoxLayout()
        add_files_btn = QPushButton("📄 Add File(s)")
        add_files_btn.clicked.connect(self._add_files)
        btn_row.addWidget(add_files_btn)

        add_dir_btn = QPushButton("📁 Add Folder")
        add_dir_btn.clicked.connect(self._add_folder)
        btn_row.addWidget(add_dir_btn)

        remove_btn = QPushButton("✖ Remove Selected")
        remove_btn.clicked.connect(self._remove_selected)
        btn_row.addWidget(remove_btn)

        clear_btn = QPushButton("🗑 Clear All")
        clear_btn.clicked.connect(self._clear_files)
        btn_row.addWidget(clear_btn)
        btn_row.addStretch()
        form_in.addLayout(btn_row)
        layout.addWidget(group_in)

        # Output & Options Group
        group_out = QGroupBox("Output Settings & Processing Options")
        form_out = QVBoxLayout(group_out)

        out_row = QHBoxLayout()
        out_row.addWidget(QLabel("Output destination:"))
        self.out_ed = QLineEdit()
        self.out_ed.setPlaceholderText("Auto-generated alongside source files if left blank")
        out_row.addWidget(self.out_ed, 1)

        browse_file_btn = QPushButton("📄 File...")
        browse_file_btn.clicked.connect(self._browse_file)
        out_row.addWidget(browse_file_btn)

        browse_folder_btn = QPushButton("📁 Folder...")
        browse_folder_btn.clicked.connect(self._browse_folder)
        out_row.addWidget(browse_folder_btn)
        form_out.addLayout(out_row)

        # Deduplication row
        dedup_row = QHBoxLayout()
        self.single_chk = QCheckBox("Single Postings only")
        self.single_chk.setChecked(False)
        self.single_chk.toggled.connect(self._on_dedup_toggled)
        dedup_row.addWidget(self.single_chk)

        dedup_row.addWidget(QLabel("Strategy:"))
        self.dedup_mode = QComboBox()
        self.dedup_mode.addItem("Prioritize POSTED (Recommended)", "prefer_posted")
        self.dedup_mode.addItem("Strictly Keep 1st Finding", "first")
        self.dedup_mode.addItem("Keep Last Finding", "last")
        self.dedup_mode.setEnabled(False)
        dedup_row.addWidget(self.dedup_mode)
        dedup_row.addStretch()
        form_out.addLayout(dedup_row)

        # Status and extra filters row
        filter_row = QHBoxLayout()
        filter_row.addWidget(QLabel("Status Filter:"))
        self.status_filter_cb = QComboBox()
        self.status_filter_cb.addItem("All Records", "all")
        self.status_filter_cb.addItem("Posted Only", "posted")
        self.status_filter_cb.addItem("Not Pos Only", "not_pos")
        filter_row.addWidget(self.status_filter_cb)

        self.file_chk = QCheckBox("Include Source File")
        filter_row.addWidget(self.file_chk)
        filter_row.addStretch()
        form_out.addLayout(filter_row)
        layout.addWidget(group_out)

        # Summary Metric Cards Row
        stats_frame = QHBoxLayout()
        stats_frame.setSpacing(8)

        self.card_files = self._create_stat_badge("Files Loaded", "0", "#3B82F6")
        stats_frame.addWidget(self.card_files)

        self.card_total = self._create_stat_badge("Total Records", "0", "#6366F1")
        stats_frame.addWidget(self.card_total)

        self.card_posted = self._create_stat_badge("Posted", "0", "#10B981")
        stats_frame.addWidget(self.card_posted)

        self.card_not_pos = self._create_stat_badge("Not Pos", "0", "#EF4444")
        stats_frame.addWidget(self.card_not_pos)

        self.card_amount = self._create_stat_badge("G. Total Sum", "Rs. 0", "#F59E0B")
        stats_frame.addWidget(self.card_amount)

        layout.addLayout(stats_frame)

        # Actions Row
        actions = QHBoxLayout()
        self.start_btn = QPushButton("▶  START PROCESSING")
        self.start_btn.setStyleSheet(START_BTN)
        self.start_btn.clicked.connect(self._start)
        actions.addWidget(self.start_btn)

        self.stop_btn = QPushButton("■  STOP")
        self.stop_btn.setStyleSheet(STOP_BTN)
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self._stop)
        actions.addWidget(self.stop_btn)

        self.open_excel_btn = QPushButton("📊 Open Excel")
        self.open_excel_btn.setEnabled(False)
        self.open_excel_btn.clicked.connect(self._open_excel)
        actions.addWidget(self.open_excel_btn)

        self.open_folder_btn = QPushButton("📂 Open Folder")
        self.open_folder_btn.setEnabled(False)
        self.open_folder_btn.clicked.connect(self._open_folder)
        actions.addWidget(self.open_folder_btn)
        actions.addStretch()
        layout.addLayout(actions)

        # Status & Progress
        self.status = QLabel("Status: Ready. Add CP-139 file(s) or folder to begin.")
        self.status.setStyleSheet("color:#4a90d9;")
        layout.addWidget(self.status)

        # Log Box
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setMinimumHeight(130)
        layout.addWidget(self.log_box)

    def _create_stat_badge(self, title: str, initial_val: str, color_hex: str) -> QFrame:
        frame = QFrame()
        frame.setStyleSheet(
            f"QFrame {{ background: #262b36; border: 1px solid #374151; border-radius: 6px; padding: 4px; }}"
        )
        box = QVBoxLayout(frame)
        box.setContentsMargins(8, 4, 8, 4)
        box.setSpacing(2)

        t_lbl = QLabel(title)
        t_lbl.setStyleSheet("color: #9ba8ba; font-size: 8pt; font-weight: 500;")
        box.addWidget(t_lbl)

        v_lbl = QLabel(initial_val)
        v_lbl.setStyleSheet(f"color: {color_hex}; font-size: 11pt; font-weight: bold;")
        box.addWidget(v_lbl)

        frame.value_label = v_lbl
        return frame

    def _on_dedup_toggled(self, checked: bool):
        self.dedup_mode.setEnabled(checked)

    def _add_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select CP-139 Posting Files",
            "",
            "Supported Reports (*.txt *.pdf);;Text Files (*.txt);;PDF Files (*.pdf);;All files (*.*)"
        )
        for p in paths:
            norm_p = os.path.normpath(p)
            if norm_p not in self.file_paths:
                self.file_paths.append(norm_p)
                self.file_list.addItem(f"📄 {os.path.basename(norm_p)}  ({norm_p})")
        self._update_default_output()
        self.card_files.value_label.setText(str(len(self.file_paths)))

    def _add_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder Containing CP-139 Files")
        if folder:
            files = sorted(glob.glob(os.path.join(folder, "*.txt")) + glob.glob(os.path.join(folder, "*.pdf")))
            for tf in files:
                norm_p = os.path.normpath(tf)
                if norm_p not in self.file_paths:
                    self.file_paths.append(norm_p)
                    self.file_list.addItem(f"📄 {os.path.basename(norm_p)}  ({norm_p})")
            self._update_default_output()
            self.card_files.value_label.setText(str(len(self.file_paths)))

    def _remove_selected(self):
        for item in self.file_list.selectedItems():
            row = self.file_list.row(item)
            self.file_list.takeItem(row)
            if row < len(self.file_paths):
                del self.file_paths[row]
        self._update_default_output()
        self.card_files.value_label.setText(str(len(self.file_paths)))

    def _clear_files(self):
        self.file_paths.clear()
        self.file_list.clear()
        self.out_ed.clear()
        self.card_files.value_label.setText("0")
        self.card_total.value_label.setText("0")
        self.card_posted.value_label.setText("0")
        self.card_not_pos.value_label.setText("0")
        self.card_amount.value_label.setText("Rs. 0")
        self.open_excel_btn.setEnabled(False)
        self.open_folder_btn.setEnabled(False)
        self.status.setText("Status: Cleared all files.")

    def _update_default_output(self):
        if self.file_paths and not self.out_ed.text().strip():
            first_dir = os.path.dirname(self.file_paths[0])
            first_base = os.path.splitext(os.path.basename(self.file_paths[0]))[0]
            date_str = datetime.now().strftime("%d%m%Y")
            self.out_ed.setText(os.path.normpath(os.path.join(first_dir, f"{first_base}_Results_{date_str}.xlsx")))
        if len(self.file_paths) > 1:
            self.file_chk.setChecked(True)

    def _browse_file(self):
        curr = self.out_ed.text().strip()
        init_dir = os.path.dirname(curr) if curr and os.path.isdir(os.path.dirname(curr)) else ""
        path, _ = QFileDialog.getSaveFileName(self, "Save Output Excel As", init_dir, "Excel Workbook (*.xlsx)")
        if path:
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            self.out_ed.setText(os.path.normpath(path))

    def _browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder:
            curr = self.out_ed.text().strip()
            name = os.path.basename(curr) if curr and curr.lower().endswith(".xlsx") else "CP139_Posting_Results.xlsx"
            self.out_ed.setText(os.path.normpath(os.path.join(folder, name)))

    def _start(self):
        if not self.file_paths:
            QMessageBox.warning(self, "No Files", "Please add at least one CP-139 report file (.txt or .pdf).")
            return

        out_path = self.out_ed.text().strip()
        if not out_path:
            self._update_default_output()
            out_path = self.out_ed.text().strip()

        self.log_box.clear()
        self.status.setText("Status: Processing CP-139 reports...")
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.open_excel_btn.setEnabled(False)
        self.open_folder_btn.setEnabled(False)

        self.worker = CP139Worker(
            file_paths=list(self.file_paths),
            output_path=out_path,
            deduplicate=self.single_chk.isChecked(),
            dedup_mode=self.dedup_mode.currentData(),
            status_filter=self.status_filter_cb.currentData(),
            include_source_file=self.file_chk.isChecked()
        )
        self.worker.log.connect(self._log)
        self.worker.progress.connect(self._on_progress)
        self.worker.done.connect(self._done)
        self.worker.failed.connect(self._failed)
        self.worker.finished.connect(lambda: (self.start_btn.setEnabled(True), self.stop_btn.setEnabled(False)))
        self.worker.start()

    def _stop(self):
        if self.worker and self.worker.isRunning():
            self.worker.requestInterruption()
            self.status.setText("Status: Stopping worker...")

    def _log(self, msg: str):
        self.log_box.append(msg)

    def _on_progress(self, current: int, total: int, message: str):
        self.status.setText(f"Status: [{current}/{total}] {message}")

    def _done(self, saved_path: str, stats: dict, records: list):
        self.last_output = saved_path
        self.status.setText(f"Status: Finished successfully. Output generated at {os.path.basename(saved_path)}")
        self.open_excel_btn.setEnabled(True)
        self.open_folder_btn.setEnabled(True)

        # Update Stat Badges
        self.card_total.value_label.setText(str(stats.get("total", 0)))
        self.card_posted.value_label.setText(str(stats.get("posted", 0)))
        self.card_not_pos.value_label.setText(str(stats.get("not_pos", 0)))
        amt = stats.get("total_amount", 0)
        self.card_amount.value_label.setText(f"Rs. {amt:,.0f}")

        msg_body = (
            f"Processed {len(self.file_paths)} report file(s).\n\n"
            f"  • Total Records in Report: {stats.get('total', 0)}\n"
            f"  • Posted: {stats.get('posted', 0)}\n"
            f"  • Not Posted: {stats.get('not_pos', 0)}\n"
            f"  • Grand Total Sum: Rs. {amt:,.2f}\n"
        )
        if stats.get("duplicates_removed", 0) > 0:
            msg_body += f"  • Duplicates Removed: {stats.get('duplicates_removed')}\n"
        msg_body += f"\nExcel file saved to:\n{saved_path}"

        notify(
            self,
            "CP 139 Extraction Complete",
            msg_body,
            file_path=saved_path,
            folder_path=os.path.dirname(saved_path)
        )

    def _failed(self, error_msg: str):
        self.status.setText("Status: Failed")
        notify(self, "CP 139 Extraction Error", error_msg, critical=True)

    def _open_excel(self):
        if self.last_output and os.path.exists(self.last_output):
            os.startfile(self.last_output)

    def _open_folder(self):
        if self.last_output:
            folder = os.path.dirname(os.path.abspath(self.last_output))
            if os.path.exists(folder):
                os.startfile(folder)


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        target_path = sys.argv[1]
        out_excel = sys.argv[2] if len(sys.argv) > 2 else os.path.splitext(target_path)[0] + "_Results.xlsx"
        print(f"Parsing: {target_path}")
        recs = parse_cp139_file(target_path)
        print(f"Extracted {len(recs)} record(s).")
        stats = export_cp139_to_excel(recs, out_excel, include_source_file=True)
        print("Excel report generated successfully:", stats)
    else:
        from PySide6.QtWidgets import QApplication
        app = QApplication(sys.argv)
        panel = CP139PostingCheckerPanel()
        panel.show()
        sys.exit(app.exec())


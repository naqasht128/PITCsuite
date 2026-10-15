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

from reportlab.lib.colors import Color, black
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from pypdf import PdfReader, PdfWriter


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


# ====================================================================
# PDF GENERATION & EXCEL MATCHING ENGINE
# ====================================================================

def download_cp139_template(save_path: str) -> None:
    """Creates a standardized template workbook for matching AC Nos against CP-139 reports."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Template"

    headers = ["AC No.", "Para Sr. No.", "Remarks"]
    hdr_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    hdr_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    hdr_align = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style="thin", color="D3D3D3"),
        right=Side(style="thin", color="D3D3D3"),
        top=Side(style="thin", color="D3D3D3"),
        bottom=Side(style="thin", color="D3D3D3")
    )

    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = hdr_align
        cell.border = thin_border

    # Sample rows
    samples = [
        ("30266111005700", "1"),
        ("09269310611170", "2"),
        ("10269310676106", "3"),
        ("14269310877502", "4"),
    ]
    for row_idx, (ac, para) in enumerate(samples, 2):
        c1 = ws.cell(row=row_idx, column=1, value=ac)
        c2 = ws.cell(row=row_idx, column=2, value=para)
        c1.alignment = Alignment(horizontal="center")
        c2.alignment = Alignment(horizontal="center")
        c1.border = thin_border
        c2.border = thin_border

    note_cell = ws.cell(row=7, column=1, value="Note: Enter the complete 14-digit Account Numbers in Col A and corresponding Para Sr. Nos in Col B. Remarks in Col C will be updated automatically.")
    note_cell.font = Font(name="Segoe UI", size=9, italic=True, color="555555")

    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 25

    os.makedirs(os.path.dirname(os.path.abspath(save_path)), exist_ok=True)
    wb.save(save_path)


def load_cp139_pages(filepath: str) -> List[Dict[str, Any]]:
    """Loads all pages from a CP-139 (.txt or .pdf) report file."""
    pages = []
    ext = os.path.splitext(filepath)[1].lower()
    fname = os.path.basename(filepath)

    if ext == ".pdf":
        try:
            reader = PdfReader(filepath)
            for idx, p in enumerate(reader.pages, start=1):
                txt = p.extract_text() or ""
                pages.append({
                    "page_num": idx,
                    "source_path": filepath,
                    "source_name": fname,
                    "lines": txt.replace("\r\n", "\n").replace("\r", "\n").split("\n"),
                    "full_text": txt,
                    "is_pdf": True,
                    "pdf_page_obj": p
                })
        except Exception:
            pass
    else:
        try:
            with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
                content = f.read()

            raw_pages = content.split('\x0c')
            if len(raw_pages) <= 1 and "DATE & TIME" in content:
                parts = re.split(r'(?m)(?=^\s*DATE & TIME\s*:)', content)
                if len(parts) > 1:
                    raw_pages = [p for p in parts if p.strip()]

            for idx, ptext in enumerate(raw_pages, start=1):
                if not ptext.strip():
                    continue
                pages.append({
                    "page_num": idx,
                    "source_path": filepath,
                    "source_name": fname,
                    "lines": ptext.replace("\r\n", "\n").replace("\r", "\n").split("\n"),
                    "full_text": ptext,
                    "is_pdf": False,
                    "pdf_page_obj": None
                })
        except Exception:
            pass

    return pages


def clean_account_number(raw_ac: Any) -> str:
    """
    Cleans and standardizes an account number from user/Excel input.
    Extracts pure digits or trimmed string.
    """
    if raw_ac is None:
        return ""
    s = str(raw_ac).strip()
    digits = re.sub(r"[^0-9]", "", s)
    return digits if digits else s


def find_account_on_page(page: Dict[str, Any], target_ac: str) -> Optional[Dict[str, Any]]:
    """Searches a page for an exact target account number."""
    if not target_ac:
        return None
    lines = page["lines"]
    for line_no, line in enumerate(lines):
        pos = line.find(target_ac)
        if pos != -1:
            before_char = line[pos - 1] if pos > 0 else " "
            after_char = line[pos + len(target_ac)] if pos + len(target_ac) < len(line) else " "
            if before_char.isdigit() or after_char.isdigit():
                continue

            grand_total = 0.0
            status_str = "Posted"

            for sub_i in range(line_no, min(line_no + 6, len(lines))):
                sub_line = lines[sub_i]
                if "Not Pos" in sub_line or "NOT POS" in sub_line or "RFP" in sub_line or "NR" in sub_line:
                    status_str = "Not Pos"
                elif "Posted" in sub_line or "POSTED" in sub_line:
                    status_str = "Posted"

                amounts = re.findall(r'(?<![0-9])-?[0-9]{1,3}(?:,[0-9]{3})*(?:\.[0-9]{2})(?![0-9])', sub_line)
                if amounts:
                    try:
                        grand_total = float(amounts[-1].replace(",", ""))
                    except ValueError:
                        pass

            return {
                "matched_cand": target_ac,
                "line_no": line_no,
                "line_text": line,
                "col_start": pos,
                "col_end": pos + len(target_ac),
                "amount": grand_total,
                "status": status_str,
                "page_num": page["page_num"],
                "source_path": page["source_path"],
                "source_name": page["source_name"]
            }
    return None


def render_highlighted_txt_page(
    page: Dict[str, Any],
    matches_on_page: List[Dict[str, Any]],
    output_pdf_path: str,
    stamp_para: bool = True,
    para_stamps: Optional[List[str]] = None
) -> None:
    """Renders a text page to landscape A4 with yellow highlight over matched lines and optional Para stamp."""
    pw, ph = landscape(A4)
    margin_x = 14
    margin_y = 16
    lines = page["lines"]
    max_len = max((len(l) for l in lines), default=80)

    font_size = min(8.0, max(4.5, (pw - 2 * margin_x) / (max_len * 0.60)))
    leading = min(9.2, max(5.0, (ph - 2 * margin_y) / max(len(lines), 1)))

    c = canvas.Canvas(output_pdf_path, pagesize=(pw, ph))
    c.setFont("Courier", font_size)

    match_lines = {m["line_no"]: m for m in matches_on_page}

    for line_no, line in enumerate(lines):
        y = ph - margin_y - (line_no + 1) * leading
        if y < 10:
            break

        if line_no in match_lines:
            m = match_lines[line_no]
            sx = margin_x + m["col_start"] * font_size * 0.60
            sw = max((m["col_end"] - m["col_start"]) * font_size * 0.60, 40)
            c.setFillColor(Color(1.0, 0.88, 0.15, alpha=0.45))
            c.rect(sx - 2, y - 2, sw + 4, leading + 2, stroke=0, fill=1)

        c.setFillColor(black)
        c.drawString(margin_x, y, line)

    if stamp_para and para_stamps:
        stamp_vals = [str(s).strip() for s in para_stamps if s and str(s).strip()]
        if stamp_vals:
            cleaned_stamps = []
            for s in stamp_vals:
                cleaned = re.sub(r'^(?:para\s*(?:sr\.?\s*(?:no\.?)?)?\s*:?\s*)', '', s, flags=re.IGNORECASE).strip()
                cleaned_stamps.append(cleaned if cleaned else s)
            joined_vals = ", ".join(cleaned_stamps)
            stamp_text = f"Para Sr. No. {joined_vals}"
            c.setFont("Helvetica-Bold", 16)
            c.setFillColor(black)
            tw = c.stringWidth(stamp_text, "Helvetica-Bold", 16)
            sx = pw - margin_x - tw - 12
            sy = ph - margin_y - 14
            c.drawString(sx, sy, stamp_text)
            c.setLineWidth(1.4)
            c.line(sx, sy - 2, sx + tw, sy - 2)

    c.save()


def render_highlighted_pdf_source_page(
    page: Dict[str, Any],
    matches_on_page: List[Dict[str, Any]],
    output_pdf_path: str,
    stamp_para: bool = True,
    para_stamps: Optional[List[str]] = None
) -> None:
    """Overlays stamp and highlights on an existing PDF page."""
    orig_page = page["pdf_page_obj"]
    pw = float(orig_page.mediabox.width)
    ph = float(orig_page.mediabox.height)

    overlay_path = output_pdf_path + ".overlay.tmp"
    c = canvas.Canvas(overlay_path, pagesize=(pw, ph))

    if stamp_para and para_stamps:
        stamp_vals = [str(s).strip() for s in para_stamps if s and str(s).strip()]
        if stamp_vals:
            cleaned_stamps = []
            for s in stamp_vals:
                cleaned = re.sub(r'^(?:para\s*(?:sr\.?\s*(?:no\.?)?)?\s*:?\s*)', '', s, flags=re.IGNORECASE).strip()
                cleaned_stamps.append(cleaned if cleaned else s)
            joined_vals = ", ".join(cleaned_stamps)
            stamp_text = f"Para Sr. No. {joined_vals}"
            c.setFont("Helvetica-Bold", 16)
            c.setFillColor(black)
            tw = c.stringWidth(stamp_text, "Helvetica-Bold", 16)
            sx = pw - 24 - tw
            sy = ph - 28
            c.drawString(sx, sy, stamp_text)
            c.setLineWidth(1.4)
            c.line(sx, sy - 2, sx + tw, sy - 2)

    c.save()
    try:
        overlay_reader = PdfReader(overlay_path)
        orig_page.merge_page(overlay_reader.pages[0])
        writer = PdfWriter()
        writer.add_page(orig_page)
        with open(output_pdf_path, "wb") as f:
            writer.write(f)
    finally:
        if os.path.exists(overlay_path):
            try:
                os.remove(overlay_path)
            except Exception:
                pass


def export_matching_results_to_excel(results: List[Dict[str, Any]], out_path: str) -> Dict[str, Any]:
    """Generates an Excel report of matched AC numbers and generated PDFs."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CP139_Matched_Results"

    headers = [
        "Sr.", "Input AC No.", "Matched Ref# in CP139",
        "Para Sr. No.", "Status", "Page No.", "Grand Total", "Posting Status",
        "Source Report", "Generated PDF File"
    ]

    hdr_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    hdr_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    border_thin = Border(
        left=Side(style="thin", color="D3D3D3"),
        right=Side(style="thin", color="D3D3D3"),
        top=Side(style="thin", color="D3D3D3"),
        bottom=Side(style="thin", color="D3D3D3")
    )
    fill_found = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid")
    fill_not_found = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    font_data = Font(name="Segoe UI", size=10)

    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = hdr_align
        cell.border = border_thin
    ws.row_dimensions[1].height = 28

    found_count = 0
    not_found_count = 0
    total_amount = 0.0

    for row_idx, r in enumerate(results, 2):
        ws.row_dimensions[row_idx].height = 20
        status = r.get("status", "Not Found")
        is_found = (status == "Found")

        if is_found:
            found_count += 1
            total_amount += float(r.get("amount", 0.0))
            row_fill = fill_found
        else:
            not_found_count += 1
            row_fill = fill_not_found

        ref_display = r.get("matched_ref", "—")

        row_values = [
            row_idx - 1,
            str(r.get("input_ac", "")),
            ref_display,
            str(r.get("para_sr", "")),
            status,
            r.get("page_num", "—"),
            r.get("amount", 0.0) if is_found else "—",
            r.get("posting_status", "—") if is_found else "—",
            r.get("source_file", "—"),
            r.get("pdf_file", "—")
        ]

        for col_idx, val in enumerate(row_values, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = font_data
            cell.border = border_thin
            cell.fill = row_fill

            if col_idx in (1, 4, 5, 6, 8):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in (2, 3):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx == 7:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                if is_found:
                    cell.number_format = "#,##0.00"
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    col_widths = [8, 20, 38, 15, 14, 12, 16, 16, 30, 35]
    for col_idx, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = w

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    wb.save(out_path)

    return {
        "total": len(results),
        "found": found_count,
        "not_found": not_found_count,
        "total_amount": total_amount,
        "saved_path": out_path
    }


class CP139Worker(QThread):
    """Background worker for parsing CP-139 posting files and generating PDFs."""
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
        include_source_file: bool = False,
        make_pdfs: bool = False,
        excel_input_path: str = "",
        single_pdf: bool = True,
        separate_pdfs: bool = False,
        stamp_para: bool = True
    ):
        super().__init__()
        self.file_paths = file_paths
        self.output_path = output_path
        self.deduplicate = deduplicate
        self.dedup_mode = dedup_mode
        self.status_filter = status_filter
        self.include_source_file = include_source_file
        self.make_pdfs = make_pdfs
        self.excel_input_path = excel_input_path
        self.single_pdf = single_pdf
        self.separate_pdfs = separate_pdfs
        self.stamp_para = stamp_para

    def run(self):
        try:
            if self.make_pdfs:
                self._run_make_pdfs()
            else:
                self._run_standard_extraction()
        except Exception as exc:
            self.failed.emit(str(exc))

    def _run_standard_extraction(self):
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

        # Resolve output destination folder and file
        if os.path.isdir(self.output_path) or not self.output_path.lower().endswith(".xlsx"):
            out_folder = self.output_path
            first_base = os.path.splitext(os.path.basename(self.file_paths[0]))[0] if self.file_paths else "CP139"
            date_str = datetime.now().strftime("%d%m%Y")
            excel_out_path = os.path.join(out_folder, f"{first_base}_Results_{date_str}.xlsx")
        else:
            excel_out_path = self.output_path
            out_folder = os.path.dirname(self.output_path)
        os.makedirs(out_folder, exist_ok=True)

        # Export to Excel
        self.log.emit(f"\n💾 Generating structured Excel report: {excel_out_path}...")
        stats = export_cp139_to_excel(
            all_records,
            excel_out_path,
            include_source_file=self.include_source_file
        )
        stats["duplicates_removed"] = dups_count
        stats["raw_total"] = total_raw
        stats["output_dir"] = out_folder
        stats["excel_path"] = excel_out_path

        self.log.emit("✔ Export completed successfully!")
        self.log.emit(f"   • Total records in Excel: {stats['total']}")
        self.log.emit(f"   • Posted: {stats['posted']}")
        self.log.emit(f"   • Not Pos: {stats['not_pos']}")
        self.log.emit(f"   • Grand Total Sum: Rs. {stats['total_amount']:,.2f}")
        self.log.emit(f"   • Output Folder: {out_folder}")
        self.log.emit(f"   • Output Excel: {os.path.basename(excel_out_path)}")

        self.done.emit(excel_out_path, stats, all_records)

    def _run_make_pdfs(self):
        self.log.emit("🚀 Starting CP-139 Make PDFs & Excel Matching...")
        if not self.excel_input_path or not os.path.isfile(self.excel_input_path):
            raise ValueError(f"Excel input file not found: {self.excel_input_path}")

        wb = openpyxl.load_workbook(self.excel_input_path)
        ws = wb.active
        excel_items = []
        for row_i in range(2, ws.max_row + 1):
            ac_val = ws.cell(row=row_i, column=1).value
            para_val = ws.cell(row=row_i, column=2).value
            if ac_val is not None and str(ac_val).strip():
                val_str = str(ac_val).strip()
                if val_str.lower().startswith("note:"):
                    continue
                excel_items.append({
                    "row_idx": row_i,
                    "input_ac": val_str,
                    "para_sr": str(para_val).strip() if para_val is not None else "",
                    "remark": "AC No Not found"
                })

        if not excel_items:
            raise ValueError("No account numbers found in the selected Excel workbook.")

        self.log.emit(f"   ↳ Loaded {len(excel_items)} account number(s) to search from Excel.")

        all_pages = []
        total_files = len(self.file_paths)
        for idx, fp in enumerate(self.file_paths, 1):
            if self.isInterruptionRequested():
                return
            self.progress.emit(idx, total_files, f"Loading pages from {os.path.basename(fp)}...")
            self.log.emit(f"📖 Loading pages from: {os.path.basename(fp)}")
            pages = load_cp139_pages(fp)
            self.log.emit(f"   ↳ Loaded {len(pages)} page(s).")
            all_pages.extend(pages)

        if not all_pages:
            raise ValueError("No pages could be loaded from the selected CP-139 file(s).")

        self.log.emit(f"\n🔍 Matching {len(excel_items)} account(s) across {len(all_pages)} report page(s)...")

        matched_results = []
        page_matches_map = {}

        for idx, item in enumerate(excel_items, 1):
            if self.isInterruptionRequested():
                return
            raw_ac = item["input_ac"]
            clean_ac = clean_account_number(raw_ac)
            match = None
            for p in all_pages:
                match = find_account_on_page(p, clean_ac)
                if match:
                    break

            if match:
                item["remark"] = "Found and PDF made"
                self.log.emit(f"   ✔ Found AC {raw_ac} on Page {match['page_num']} ({match['source_name']}) -> Ref# {match['matched_cand']}")
                key = (match["source_path"], match["page_num"])
                if key not in page_matches_map:
                    target_p = next((p for p in all_pages if p["source_path"] == match["source_path"] and p["page_num"] == match["page_num"]), None)
                    page_matches_map[key] = {
                        "page": target_p,
                        "matches": [],
                        "para_stamps": []
                    }
                match["input_ac"] = raw_ac
                match["para_sr"] = item["para_sr"]
                page_matches_map[key]["matches"].append(match)
                if item["para_sr"] and item["para_sr"] not in page_matches_map[key]["para_stamps"]:
                    page_matches_map[key]["para_stamps"].append(item["para_sr"])

                matched_results.append({
                    "input_ac": raw_ac,
                    "matched_ref": match["matched_cand"],
                    "para_sr": item["para_sr"],
                    "status": "Found",
                    "page_num": match["page_num"],
                    "amount": match["amount"],
                    "posting_status": match["status"],
                    "source_file": match["source_name"],
                    "pdf_file": ""
                })
            else:
                item["remark"] = "AC No Not found"
                self.log.emit(f"   ✖ Not Found: AC {raw_ac}")
                matched_results.append({
                    "input_ac": raw_ac,
                    "matched_ref": "—",
                    "para_sr": item["para_sr"],
                    "status": "Not Found",
                    "page_num": "—",
                    "amount": 0.0,
                    "posting_status": "—",
                    "source_file": "—",
                    "pdf_file": "—"
                })

        # Resolve output destination folder and file
        if os.path.isdir(self.output_path) or not self.output_path.lower().endswith(".xlsx"):
            out_folder = self.output_path
            first_base = os.path.splitext(os.path.basename(self.file_paths[0]))[0] if self.file_paths else "CP139"
            date_str = datetime.now().strftime("%d%m%Y")
            excel_out_path = os.path.join(out_folder, f"{first_base}_Results_{date_str}.xlsx")
        else:
            excel_out_path = self.output_path
            out_folder = os.path.dirname(self.output_path)

        pdf_out_dir = out_folder
        os.makedirs(pdf_out_dir, exist_ok=True)

        self.log.emit(f"\n🎨 Rendering {len(page_matches_map)} highlighted PDF page(s)...")

        rendered_temp_pdfs = []
        for (src_path, p_num), p_data in page_matches_map.items():
            if self.isInterruptionRequested():
                return
            p_obj = p_data["page"]
            stamps = p_data["para_stamps"]
            matches = p_data["matches"]

            # Naming logic:
            # In separate pdf output: files name are on AC Nos, and if Para Sr. No. is also checked then File names are on Para Sr. Nos.
            page_paras = [m.get("para_sr", "").strip() for m in matches if m.get("para_sr", "").strip()]
            clean_paras = []
            for p in page_paras:
                cp = re.sub(r'^(?:para\s*(?:sr\.?\s*(?:no\.?)?)?\s*:?\s*)', '', p, flags=re.IGNORECASE).strip()
                clean_paras.append(cp if cp else p)

            page_acs = [m.get("input_ac", "").strip() or m.get("matched_cand", "").strip() for m in matches if (m.get("input_ac") or m.get("matched_cand"))]

            if self.stamp_para and clean_paras:
                base_title = "_".join(dict.fromkeys(clean_paras))
            elif page_acs:
                base_title = "_".join(dict.fromkeys(page_acs))
            else:
                base_title = f"CP139_{Path(src_path).stem}_Page{p_num}"

            safe_title = re.sub(r'[\\/*?:"<>|]', '_', base_title).strip()
            single_out_name = f"{safe_title}.pdf"
            single_out_path = os.path.join(pdf_out_dir, single_out_name)

            counter = 1
            while single_out_path in rendered_temp_pdfs:
                single_out_name = f"{safe_title}_{counter}.pdf"
                single_out_path = os.path.join(pdf_out_dir, single_out_name)
                counter += 1

            if p_obj["is_pdf"]:
                render_highlighted_pdf_source_page(p_obj, p_data["matches"], single_out_path, stamp_para=self.stamp_para, para_stamps=stamps)
            else:
                render_highlighted_txt_page(p_obj, p_data["matches"], single_out_path, stamp_para=self.stamp_para, para_stamps=stamps)

            rendered_temp_pdfs.append(single_out_path)

            for res in matched_results:
                if res["status"] == "Found" and res["source_file"] == os.path.basename(src_path) and res["page_num"] == p_num:
                    res["pdf_file"] = single_out_name

        final_pdf_path = ""
        if self.single_pdf and rendered_temp_pdfs:
            self.log.emit("📑 Combining highlighted pages into a Single PDF...")
            combined_name = f"CP139_Matched_Postings_Combined_{datetime.now().strftime('%d%m%Y_%H%M%S')}.pdf"
            combined_path = os.path.join(pdf_out_dir, combined_name)
            writer = PdfWriter()
            for tmp_p in rendered_temp_pdfs:
                r = PdfReader(tmp_p)
                for page in r.pages:
                    writer.add_page(page)
            with open(combined_path, "wb") as f:
                writer.write(f)

            final_pdf_path = combined_path
            self.log.emit(f"✔ Single PDF generated successfully: {combined_name}")

            if not self.separate_pdfs:
                for tmp_p in rendered_temp_pdfs:
                    try:
                        os.remove(tmp_p)
                    except Exception:
                        pass
                for res in matched_results:
                    if res["status"] == "Found":
                        res["pdf_file"] = combined_name
            else:
                self.log.emit(f"✔ Retained {len(rendered_temp_pdfs)} separate PDF file(s).")
        elif self.separate_pdfs and rendered_temp_pdfs:
            final_pdf_path = rendered_temp_pdfs[0]
            self.log.emit(f"✔ Generated {len(rendered_temp_pdfs)} separate PDF file(s) in: {pdf_out_dir}")

        # Update Column C ('Remarks') in the user's Input Excel workbook
        self.log.emit(f"\n📝 Updating Remarks in Input Excel: {os.path.basename(self.excel_input_path)}...")
        try:
            hdr_cell = ws.cell(row=1, column=3, value="Remarks")
            hdr_cell.font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
            hdr_cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
            hdr_cell.alignment = Alignment(horizontal="center", vertical="center")
            border_thin = Border(
                left=Side(style="thin", color="D3D3D3"),
                right=Side(style="thin", color="D3D3D3"),
                top=Side(style="thin", color="D3D3D3"),
                bottom=Side(style="thin", color="D3D3D3")
            )
            hdr_cell.border = border_thin
            ws.row_dimensions[1].height = 28

            fill_found = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid")
            fill_not_found = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
            font_data = Font(name="Segoe UI", size=10)

            for item in excel_items:
                r_idx = item["row_idx"]
                c_val = item.get("remark", "AC No Not found")
                c_cell = ws.cell(row=r_idx, column=3, value=c_val)
                c_cell.font = font_data
                c_cell.alignment = Alignment(horizontal="center", vertical="center")
                c_cell.border = border_thin
                c_cell.fill = fill_found if c_val == "Found and PDF made" else fill_not_found
                ws.row_dimensions[r_idx].height = 20

            ws.column_dimensions["C"].width = 25
            wb.save(self.excel_input_path)
            self.log.emit(f"✔ Remarks updated in Input Excel (Col C): {os.path.basename(self.excel_input_path)}")
        except PermissionError:
            self.log.emit(f"⚠ Notice: Could not write remarks directly into '{os.path.basename(self.excel_input_path)}' because the file is open in Microsoft Excel. Please close it if you wish to update remarks in-place.")
        except Exception as exc:
            self.log.emit(f"⚠ Notice: Unable to update remarks in input Excel file: {exc}")

        self.log.emit(f"\n💾 Generating matching results Excel: {excel_out_path}...")
        stats = export_matching_results_to_excel(matched_results, excel_out_path)
        stats["output_dir"] = pdf_out_dir
        stats["excel_path"] = excel_out_path
        stats["pdf_path"] = final_pdf_path or pdf_out_dir
        stats["pdf_count"] = len(page_matches_map)

        self.log.emit("✔ Make PDFs & Excel matching complete!")
        self.log.emit(f"   • Total ACs processed: {stats['total']}")
        self.log.emit(f"   • Found: {stats['found']}")
        self.log.emit(f"   • Not Found: {stats['not_found']}")
        self.log.emit(f"   • Grand Total Sum: Rs. {stats['total_amount']:,.2f}")
        self.log.emit(f"   • Output Folder: {pdf_out_dir}")
        self.log.emit(f"   • Output Excel: {os.path.basename(excel_out_path)}")
        if final_pdf_path:
            self.log.emit(f"   • Output PDF: {os.path.basename(final_pdf_path)}")

        self.done.emit(excel_out_path, stats, matched_results)


class CP139PostingCheckerPanel(QWidget):
    """UI Panel for CP 139 Posting Check/PDF in PITC Suite."""
    def __init__(self):
        super().__init__()
        self.worker = None
        self.file_paths = []
        self.last_output = ""

        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        # Header Title and Description
        layout.addWidget(lbl("CP 139 Posting Check/PDF", bold=True, color="#7eb8f7"))
        layout.addWidget(lbl(
            "Extract Ref#, S/Div, Grand Total, match against Excel AC numbers, highlight records, and generate PDF pages.",
            color="#556"
        ))
        layout.addWidget(hline())

        # Section 1: Source Reports Group
        group_in = QGroupBox("Source Reports")
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

        # Section 2: Output Settings Group
        group_out = QGroupBox("Output Settings")
        form_out = QVBoxLayout(group_out)

        out_row = QHBoxLayout()
        out_row.addWidget(QLabel("Output Folder:"))
        self.out_dir_ed = QLineEdit()
        self.out_dir_ed.setPlaceholderText("Auto-generated alongside source files if left blank")
        out_row.addWidget(self.out_dir_ed, 1)

        browse_folder_btn = QPushButton("📁 Browse Folder...")
        browse_folder_btn.clicked.connect(self._browse_folder)
        out_row.addWidget(browse_folder_btn)
        form_out.addLayout(out_row)

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

        # Section 3: PDF Generation Group
        group_pdf = QGroupBox("PDF Generation")
        form_pdf = QVBoxLayout(group_pdf)

        chk_row = QHBoxLayout()
        self.make_pdf_chk = QCheckBox("Make PDFs")
        self.make_pdf_chk.setChecked(False)
        self.make_pdf_chk.toggled.connect(self._on_make_pdf_toggled)
        chk_row.addWidget(self.make_pdf_chk)

        self.single_pdf_chk = QCheckBox("Single PDF")
        self.single_pdf_chk.setChecked(True)
        self.single_pdf_chk.setEnabled(False)
        self.single_pdf_chk.toggled.connect(self._on_single_pdf_toggled)
        chk_row.addWidget(self.single_pdf_chk)

        self.separate_pdf_chk = QCheckBox("Separate PDFs")
        self.separate_pdf_chk.setChecked(False)
        self.separate_pdf_chk.setEnabled(False)
        self.separate_pdf_chk.toggled.connect(self._on_separate_pdf_toggled)
        chk_row.addWidget(self.separate_pdf_chk)

        self.stamp_para_chk = QCheckBox("Para No Stamped")
        self.stamp_para_chk.setChecked(True)
        self.stamp_para_chk.setEnabled(False)
        chk_row.addWidget(self.stamp_para_chk)
        chk_row.addStretch()
        form_pdf.addLayout(chk_row)

        excel_row = QHBoxLayout()
        excel_row.addWidget(QLabel("Excel input (.xlsx):"))
        self.excel_in_ed = QLineEdit()
        self.excel_in_ed.setPlaceholderText("Select Excel workbook containing Col A: AC No., Col B: Para Sr. No.")
        self.excel_in_ed.setEnabled(False)
        excel_row.addWidget(self.excel_in_ed, 1)

        self.browse_excel_btn = QPushButton("📄 Browse Excel...")
        self.browse_excel_btn.setEnabled(False)
        self.browse_excel_btn.clicked.connect(self._browse_excel)
        excel_row.addWidget(self.browse_excel_btn)

        self.download_tmpl_btn = QPushButton("📥 Download Template")
        self.download_tmpl_btn.setEnabled(False)
        self.download_tmpl_btn.clicked.connect(self._download_template)
        excel_row.addWidget(self.download_tmpl_btn)
        form_pdf.addLayout(excel_row)

        layout.addWidget(group_pdf)

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
        actions.addStretch()
        layout.addLayout(actions)

        # Section 4: Logs Group
        group_logs = QGroupBox("Logs")
        form_logs = QVBoxLayout(group_logs)

        self.status = QLabel("Status: Ready. Add CP-139 file(s) or folder to begin.")
        self.status.setStyleSheet("color:#4a90d9;")
        form_logs.addWidget(self.status)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setMinimumHeight(150)
        form_logs.addWidget(self.log_box)

        layout.addWidget(group_logs)

    def _on_single_pdf_toggled(self, checked: bool):
        if not self.make_pdf_chk.isChecked():
            return
        if checked:
            self.separate_pdf_chk.blockSignals(True)
            self.separate_pdf_chk.setChecked(False)
            self.separate_pdf_chk.blockSignals(False)
            self.separate_pdf_chk.setEnabled(False)
        else:
            self.separate_pdf_chk.setEnabled(True)

    def _on_separate_pdf_toggled(self, checked: bool):
        if not self.make_pdf_chk.isChecked():
            return
        if checked:
            self.single_pdf_chk.blockSignals(True)
            self.single_pdf_chk.setChecked(False)
            self.single_pdf_chk.blockSignals(False)
            self.single_pdf_chk.setEnabled(False)
        else:
            self.single_pdf_chk.setEnabled(True)

    def _on_make_pdf_toggled(self, checked: bool):
        self.stamp_para_chk.setEnabled(checked)
        self.excel_in_ed.setEnabled(checked)
        self.browse_excel_btn.setEnabled(checked)
        self.download_tmpl_btn.setEnabled(checked)

        if checked:
            if self.separate_pdf_chk.isChecked():
                self.separate_pdf_chk.setEnabled(True)
                self.single_pdf_chk.setEnabled(False)
            else:
                self.single_pdf_chk.setEnabled(True)
                self.single_pdf_chk.setChecked(True)
                self.separate_pdf_chk.setChecked(False)
                self.separate_pdf_chk.setEnabled(False)
        else:
            self.single_pdf_chk.setEnabled(False)
            self.separate_pdf_chk.setEnabled(False)

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

    def _remove_selected(self):
        for item in self.file_list.selectedItems():
            row = self.file_list.row(item)
            self.file_list.takeItem(row)
            if row < len(self.file_paths):
                del self.file_paths[row]
        self._update_default_output()

    def _clear_files(self):
        self.file_paths.clear()
        self.file_list.clear()
        self.out_dir_ed.clear()
        self.status.setText("Status: Cleared all files.")

    def _update_default_output(self):
        if self.file_paths and not self.out_dir_ed.text().strip():
            first_dir = os.path.dirname(self.file_paths[0])
            self.out_dir_ed.setText(os.path.normpath(first_dir))
        if len(self.file_paths) > 1:
            self.file_chk.setChecked(True)

    def _browse_excel(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Excel Input File",
            "",
            "Excel Files (*.xlsx *.xls);;All Files (*.*)"
        )
        if path:
            self.excel_in_ed.setText(os.path.normpath(path))

    def _download_template(self):
        curr_dir = self.out_dir_ed.text().strip() or (os.path.dirname(self.file_paths[0]) if self.file_paths else "")
        default_file = os.path.join(curr_dir, "CP139_Posting_Template.xlsx") if curr_dir else "CP139_Posting_Template.xlsx"
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Posting Template As",
            default_file,
            "Excel Workbook (*.xlsx);;All files (*.*)"
        )
        if path:
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            download_cp139_template(path)
            box = QMessageBox(QMessageBox.Information, "Template Saved", f"CP-139 template saved to:\n{path}", parent=self)
            box.addButton("Open File", QMessageBox.ActionRole).clicked.connect(lambda: os.startfile(path))
            box.addButton(QMessageBox.Ok)
            box.open()

    def _browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder", self.out_dir_ed.text().strip())
        if folder:
            self.out_dir_ed.setText(os.path.normpath(folder))

    def _start(self):
        if not self.file_paths:
            QMessageBox.warning(self, "No Files", "Please add at least one CP-139 report file (.txt or .pdf).")
            return

        is_make_pdf = self.make_pdf_chk.isChecked()
        if is_make_pdf:
            excel_in = self.excel_in_ed.text().strip()
            if not excel_in or not os.path.isfile(excel_in):
                QMessageBox.warning(self, "Missing Excel Input", "Please select an existing Excel workbook containing AC numbers to match.")
                return
            if not self.single_pdf_chk.isChecked() and not self.separate_pdf_chk.isChecked():
                QMessageBox.warning(self, "PDF Option Required", "Please select either 'Single PDF' or 'Separate PDFs'.")
                return

        out_dir = self.out_dir_ed.text().strip()
        if not out_dir:
            self._update_default_output()
            out_dir = self.out_dir_ed.text().strip()

        os.makedirs(out_dir, exist_ok=True)

        self.log_box.clear()
        self.status.setText("Status: Processing CP-139 reports...")
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)

        self.worker = CP139Worker(
            file_paths=list(self.file_paths),
            output_path=out_dir,
            deduplicate=self.single_chk.isChecked(),
            dedup_mode=self.dedup_mode.currentData(),
            status_filter=self.status_filter_cb.currentData(),
            include_source_file=self.file_chk.isChecked(),
            make_pdfs=is_make_pdf,
            excel_input_path=self.excel_in_ed.text().strip() if is_make_pdf else "",
            single_pdf=self.single_pdf_chk.isChecked() if is_make_pdf else False,
            separate_pdfs=self.separate_pdf_chk.isChecked() if is_make_pdf else False,
            stamp_para=self.stamp_para_chk.isChecked() if is_make_pdf else False
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
        out_folder = stats.get("output_dir") or os.path.dirname(saved_path)
        self.status.setText(f"Status: Finished successfully. Output in {os.path.basename(out_folder)}")

        amt = stats.get("total_amount", 0.0)
        if self.make_pdf_chk.isChecked():
            msg_body = (
                f"CP 139 Processing & PDF Generation Complete!\n\n"
                f"  • Total ACs Searched: {stats.get('total', 0)}\n"
                f"  • Found & Highlighted: {stats.get('found', 0)}\n"
                f"  • Not Found: {stats.get('not_found', 0)}\n"
                f"  • Grand Total Sum: Rs. {amt:,.2f}\n"
                f"  • PDF Pages Processed: {stats.get('pdf_count', 0)}\n\n"
                f"Output Folder:\n{out_folder}\n\n"
                f"Excel Results File:\n{os.path.basename(saved_path)}"
            )
        else:
            msg_body = (
                f"Processed {len(self.file_paths)} report file(s).\n\n"
                f"  • Total Records in Report: {stats.get('total', 0)}\n"
                f"  • Posted: {stats.get('posted', 0)}\n"
                f"  • Not Posted: {stats.get('not_pos', 0)}\n"
                f"  • Grand Total Sum: Rs. {amt:,.2f}\n"
            )
            if stats.get("duplicates_removed", 0) > 0:
                msg_body += f"  • Duplicates Removed: {stats.get('duplicates_removed')}\n"
            msg_body += f"\nOutput Folder:\n{out_folder}\n\nExcel Results File:\n{os.path.basename(saved_path)}"

        # Dialog with Open Excel and Open Folder options
        box = QMessageBox(QMessageBox.Information, "CP 139 Completed", msg_body, parent=self)
        box.setAttribute(Qt.WA_DeleteOnClose)

        open_excel_btn = box.addButton("Open Excel", QMessageBox.ActionRole)
        open_excel_btn.clicked.connect(lambda: os.startfile(saved_path) if os.path.exists(saved_path) else None)

        open_folder_btn = box.addButton("Open Folder", QMessageBox.ActionRole)
        open_folder_btn.clicked.connect(lambda: os.startfile(out_folder) if os.path.exists(out_folder) else None)

        box.addButton(QMessageBox.Ok)
        box.open()

    def _failed(self, error_msg: str):
        self.status.setText("Status: Failed")
        notify(self, "CP 139 Error", error_msg, critical=True)


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

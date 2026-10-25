"""Desktop PITC Bill Scraper & MDI Analyzer.

The selected .xlsx or .ods workbook is updated in place using its existing headings.
Excluded fields are removed before requests are made, and returned values are
written strictly below their matching headings without column shifting.
"""

from __future__ import annotations

import io
import os
import queue
import re
import subprocess
import sys
import threading
import time
import tkinter as tk
import xml.etree.ElementTree as etree
import zipfile
from decimal import Decimal
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook

URL = "https://bill.pitc.com.pk/pescobill"


def resource_path(relative_path):
    """Resolve bundled resources in both source runs and PyInstaller builds."""
    bundle_root = getattr(sys, "_MEIPASS", Path(__file__).resolve().parent)
    return str(Path(bundle_root) / relative_path)


# Fields that are intentionally never requested/exported.
DROP_LETTERS = {"B", "C", "E", "H", "K", "L", "M", "N", "P", "Q", "S"}
DROP_RANGES = ((24, 50), (28, 55), (36, 64), (67, 67), (98, 214), (217, 244))  # X:AX, AB:BC, AJ:BL, BO, CT:HF and HI:IJ
DROP_EXACT = {
    "reference_no", "consumer_id", "sub_division", "tariff_cat", "transformer",
    "energy_charges", "taxes", "arrears", "installment",
    "tracking_id",
    "total_electricity_charges", "net_electricity_charges", "subsidies", "category", "26",
    "kwh_units_44", "kwh_units_45", "kwh_units_46",
    "table_1_kwh_type_1",
    "table_1_kvarh_meter_1", "table_1_kvarh_meter_2", "table_1_kvarh_meter_3", "table_1_kvarh_meter_4",
    "table_1_kvarh_type_1", "table_1_mdi_meter_1", "table_1_mdi_meter_2", "table_1_mdi_meter_3", "table_1_mdi_meter_4",
    "table_1_mdi_type_1", "table_1_mdi_previous_1",
    "table_3_787_1", "table_3_787_off_peak_1", "table_3_787_peak_1",
    "table_3_2402_1", "table_3_2402_off_peak_1", "table_3_2402_peak_1",
    "table_3_1615_1", "table_3_1615_off_peak_1", "table_3_1615_peak_1",
}


def excel_col(n):
    result = ""
    while n:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def excel_numeric_value(value):
    """Convert bill numbers such as '41,839' to values Excel treats as numeric."""
    if value is None or isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value
    text = clean(value).replace(",", "")
    text = re.sub(r"\s*(?:kw|kva)\s*$", "", text, flags=re.I)
    if not re.fullmatch(r"[-+]?(?:\d+(?:\.\d+)?|\.\d+)", text):
        return value
    number = float(text)
    return int(number) if number.is_integer() else number


def should_be_numeric(key):
    """Identify bill fields that should be written as numbers, not text."""
    normalized = normalized_key(key).lower()
    if normalized in {
        "payment", "total payment", "total_payment", "total units", "total_units", "grand_total", "adjustments", "san_load",
        "current bill", "current_bill", "total fpa", "total_fpa",
        "prv_cum_mdi_1", "prv_cum_mdi_2", "cum_mdi_1", "cum_mdi_2",
    }:
        return True
    if re.fullmatch(r"(?:kwh|kvarh|mdi)_(?:present|read|mf|units)(?:_\d+)?", normalized):
        return True
    fkey = field_key(key)
    if re.fullmatch(r"(?:units|bill_rs|payment_rs)_\d+", fkey):
        return True
    return False


def clean(value):
    return re.sub(r"\s+", " ", str(value) if value is not None else "").strip()


def field_key(value):
    return re.sub(r"[^a-zA-Z0-9]+", "_", clean(value).lower()).strip("_") or "field"


def parse_num(val):
    if val is None:
        return None
    cleaned = re.sub(r"[^\d.]", "", str(val))
    try:
        return float(cleaned) if cleaned else None
    except ValueError:
        return None


def compute_protected_status(tariff, curr_units, hist_units):
    """Determine if a consumer is Protected or Non-Protected:
    If consumer's tariff is A-1A(01) and he has not consumed above 200 units in last 6 months
    (current month + 5 months from history), put 'Protected', otherwise put 'Non-Protected'.
    For any tariff other than A-1A(01), it remains blank."""
    if not tariff:
        return ""
    clean_tariff = re.sub(r"\s+", "", str(tariff)).upper()
    if clean_tariff != "A-1A(01)":
        return ""
    if curr_units is not None and curr_units > 200:
        return "Non-Protected"
    for u in hist_units:
        if u is not None and u > 200:
            return "Non-Protected"
    return "Protected"


def compute_total_payment(record):
    """Sum of all payments which are already fetched in excel."""
    rem = str(record.get("Remarks") or record.get("scrape_status", "")).strip().lower()
    if rem in ("bill not found", "error"):
        return ""
    total = 0.0
    has_pay = False
    for k, val in record.items():
        if k.startswith("_"):
            continue
        if k.lower() == "payment" or re.match(r"^payment\s*(?:\(rs\.\))?_\d+$", k, re.I):
            p = parse_num(val)
            if p is not None:
                total += p
                has_pay = True
    if not has_pay:
        return 0
    return int(total) if total.is_integer() else round(total, 2)


def compute_total_units(record):
    """Sum of all units consumed fetched in excel (current month kWh units + history Units)."""
    rem = str(record.get("Remarks") or record.get("scrape_status", "")).strip().lower()
    if rem in ("bill not found", "error"):
        return ""
    total = 0.0
    has_units = False

    # 1. Current month units (kwh_units_1 .. 4, or kwh_units)
    curr_sum = 0.0
    has_curr = False
    for i in range(1, 5):
        val = parse_num(record.get(f"kwh_units_{i}"))
        if val is not None:
            curr_sum += val
            has_curr = True
    if not has_curr and record.get("kwh_units") is not None:
        val = parse_num(record.get("kwh_units"))
        if val is not None:
            curr_sum += val
            has_curr = True

    if has_curr:
        total += curr_sum
        has_units = True

    # 2. Historical units (Units_1 .. Units_12)
    for k, val in record.items():
        if k.startswith("_"):
            continue
        if re.match(r"^units?_\d+$", k, re.I):
            u = parse_num(val)
            if u is not None:
                total += u
                has_units = True

    if not has_units:
        return 0
    return int(total) if total.is_integer() else round(total, 2)


def extract_meter_type_and_no(raw_meter):
    """Extract meter type prefix (e.g. S-P, 3-P) from meter number."""
    if not raw_meter:
        return "", ""
    s = clean(raw_meter)
    m = re.match(r"^([A-Za-z0-9]+(?:[-/][A-Za-z0-9]+)?)\s+(\S.*)$", s)
    if m:
        prefix = m.group(1).strip()
        if re.search(r"[A-Za-z\-]", prefix):
            meter_no = m.group(2).strip()
            return prefix, meter_no
    return "", s


def normalized_key(value):
    key = field_key(value)
    renamed = {
        "kwh_units_30": "kvarh_units_1",
        "kwh_units_31": "mdi_read_1", "kwh_units_32": "mdi_read_2",
        "kwh_units_33": "mdi_read_3", "kwh_units_34": "mdi_read_4",
        "kwh_units_35": "mdi_mf_1", "kwh_units_36": "mdi_mf_2",
        "kwh_units_37": "mdi_mf_3", "kwh_units_38": "mdi_mf_4",
        "kwh_units_39": "mdi_units_1",
        "kwh_units_40": "prv_cum_mdi_1", "kwh_units_41": "prv_cum_mdi_2",
        "kwh_units_42": "cum_mdi_1", "kwh_units_43": "cum_mdi_2",
        "mdi_present_1": "mdi_read_1", "mdi_present_2": "mdi_read_2",
        "mdi_present_3": "mdi_read_3", "mdi_present_4": "mdi_read_4",
        "prv_cumm_mdi1": "prv_cum_mdi_1", "prv_cumm_mdi2": "prv_cum_mdi_2",
        "prv_cumm_mdi_1": "prv_cum_mdi_1", "prv_cumm_mdi_2": "prv_cum_mdi_2",
        "prv_cum_mdi1": "prv_cum_mdi_1", "prv_cum_mdi2": "prv_cum_mdi_2",
        "previous_cumm_mdi1": "prv_cum_mdi_1", "previous_cumm_mdi2": "prv_cum_mdi_2",
        "previous_cumm_mdi_1": "prv_cum_mdi_1", "previous_cumm_mdi_2": "prv_cum_mdi_2",
        "previous_cum_mdi1": "prv_cum_mdi_1", "previous_cum_mdi2": "prv_cum_mdi_2",
        "previous_cum_mdi_1": "prv_cum_mdi_1", "previous_cum_mdi_2": "prv_cum_mdi_2",
        "cum_mdi1": "cum_mdi_1", "cum_mdi2": "cum_mdi_2",
        "prs_cum_mdi1": "cum_mdi_1", "prs_cum_mdi2": "cum_mdi_2",
        "prs_cumm_mdi1": "cum_mdi_1", "prs_cumm_mdi2": "cum_mdi_2",
        "diff_cum_mdi_1": "Diff_cum_mdi1", "diff_cum_mdi1": "Diff_cum_mdi1",
        "diff_cum_mdi_2": "diff_cum_mdi2", "diff_cum_mdi2": "diff_cum_mdi2",
        "actual_mdi": "Actual MDI",
        "actual_fix_charges": "Actual_Fix_Charges",
        "in_bill_fix_charges": "In_Bill_Fix_Charges (Estimated)",
        "in_bill_fix_charges_estimated": "In_Bill_Fix_Charges (Estimated)",
        "current_bill": "Current Bill",
        "current bill": "Current Bill",
        "total_fpa": "Total FPA",
        "total fpa": "Total FPA",
        "kvarh_units": "kvarh_units_1",
        "mdi_units": "mdi_units_1",
        "scrape_status": "Remarks",
        "remarks": "Remarks",
        "bno": "BNo",
        "batch_no": "BNo",
        "sdiv": "SDiv",
        "s_div": "SDiv",
        "meter_type": "Meter Type",
        "metertype": "Meter Type",
        "total_payment": "Total Payment",
        "total payment": "Total Payment",
        "total_units": "Total Units",
        "total units": "Total Units",
        "protected_non_protected": "Protected/Non-Protected",
        "protected/non-protected": "Protected/Non-Protected",
    }
    if key in renamed:
        return renamed[key]
    if key == "table_3_787_mdi_1":
        return "export_mdi"
    if key == "table_3_2402_mdi_1":
        return "import_mdi"
    m = re.match(r"^(month|status|units?|bill(?:_rs)?|payment(?:_rs)?)_?(\d+)$", key)
    if m:
        t, num = m.group(1), m.group(2)
        if t == "month":
            return f"Month_{num}"
        if t == "status":
            return f"Status_{num}"
        if t in ("units", "unit"):
            return f"Units_{num}"
        if t.startswith("bill"):
            return f"Bill (Rs.)_{num}"
        if t.startswith("payment"):
            return f"Payment (Rs.)_{num}"
    m2 = re.match(r"^(bill|payment)_(\d+)_rs$", key)
    if m2:
        t, num = m2.group(1), m2.group(2)
        return f"{t.capitalize()} (Rs.)_{num}"
    key = re.sub(r"^table_\d+_", "", key)
    return re.sub(r"^\d+(?:_\d+)*_", "", key)


def is_preserved_meter_key(key):
    k_lower = key.lower()
    return (
        k_lower.startswith("kvarh_present_")
        or k_lower.startswith("kvarh_units")
        or k_lower.startswith("kwh_units")
        or k_lower.startswith("mdi_read_")
        or k_lower.startswith("mdi_mf_")
        or k_lower.startswith("mdi_units")
        or k_lower.startswith("prv_cum_mdi_")
        or k_lower.startswith("cum_mdi_")
        or k_lower == "status"
        or k_lower in {"bno", "sdiv", "meter type", "meter_type", "protected/non-protected", "protected_non_protected", "total payment", "total_payment", "total units", "total_units"}
        or bool(re.match(r"^(?:month|status|units|bill|payment)(?:_|\s*)\d+", k_lower))
        or k_lower.startswith("in_bill_fix_charges")
        or k_lower in {
            "diff_cum_mdi1", "diff_cum_mdi2",
            "actual mdi", "actual_mdi",
            "actual_fix_charges",
            "current bill", "current_bill",
            "total fpa", "total_fpa",
        }
    )


def is_disc_status(status_str):
    """Check if status is Disconnected (P.Disc or T.Disc), where minus and actual fixed charges are skipped."""
    if not status_str:
        return False
    s = str(status_str).lower().strip()
    if re.search(r"\bp\.?\s*disc", s) or "p-disc" in s or "pdisc" in s or "permanent" in s:
        return True
    if re.search(r"\bt\.?\s*disc", s) or "t-disc" in s or "tdisc" in s or "temporary" in s:
        return True
    return False


def is_replaced_status(status_str):
    """Check if status is Replaced."""
    if not status_str:
        return False
    s = str(status_str).lower().strip()
    return "replace" in s or re.search(r"\brep\b", s)


def is_dropped(name, source_col=None, legacy_columns=False):
    key = field_key(name)
    normalized = normalized_key(name)
    if is_preserved_meter_key(normalized) or is_preserved_meter_key(key):
        return False
    if key in {"remarks", "scrape_status"} or normalized == "Remarks":
        return False
    if source_col is not None:
        in_dropped_range = any(start <= source_col <= end for start, end in DROP_RANGES)
        if (legacy_columns and excel_col(source_col) in DROP_LETTERS) or in_dropped_range:
            return True
    match = re.fullmatch(r"kwh_units_(\d+)", key)
    if match and int(match.group(1)) > 4:
        return True
    return (
        key in DROP_EXACT
        or "previous" in key
        or "bill_history" in key
        or key.startswith(("kvarh_meter_", "kvarh_mf_", "mdi_meter_"))
        or key.startswith("bill_history")
        or key.startswith(("table_4_", "table_5_", "table_6_"))
        or key in {"account_no", "raw_bill_text"}
    )


GENERAL_LABELS = {
    "san load": "san_load",
    "san load kw": "san_load",
    "tariff": "tariff",
    "status": "status",
    "feeder": "feeder",
    "name address": "name_address",
    "name and address": "name_address",
    "energy charges": "energy_charges",
    "taxes": "taxes",
    "current bill": "Current Bill",
    "current bill amount": "Current Bill",
    "total fpa": "Total FPA",
    "total f p a": "Total FPA",
    "arrears": "arrears",
    "installment": "installment",
    "adjustment": "adjustments",
    "adjustments": "adjustments",
    "grand total": "grand_total",
    "amount paid": "payment",
    "paid amount": "payment",
    "payment": "payment",
    "payment date": "payment_date",
    "ntn no": "ntn_no",
    "bill month": "bill_month",
    "due date": "due_date",
    "transformer": "transformer",
    "tariff category": "tariff_cat",
    "curr mdi": "curr_mdi",
    "sub division": "sub_division",
}

BILL_HEADING_KEYWORDS = {
    "reference no", "consumer id", "name address", "name and address", "transformer",
    "feeder", "sub division", "category", "status", "tariff category", "san load",
    "san load kw", "curr mdi", "tariff", "meter no", "meter number", "meter", "mf",
    "previous reading", "present reading", "units", "unit", "bill month", "due date",
    "total electricity charges", "subsidies", "net electricity charges", "taxes",
    "current bill", "current bill amount", "total fpa", "total f p a", "total fuel price",
    "arrears", "installment", "adjustment", "adjustments", "grand total",
    "amount paid", "paid amount", "payment", "payment date", "ntn no",
    "bill charges breakdown", "meter info", "meter information", "bill history",
    "payment history", "billing payment history", "month", "bill",
}


def canonical_label(value):
    label = re.sub(r"[^a-z0-9]+", " ", clean(value).lower()).strip()
    return GENERAL_LABELS.get(label, normalized_key(value))


def order_output_headers(headers):
    """Keep tariff/san-load stable and keep meter readings under fixed names."""
    first = [
        "AC No.", "BNo", "SDiv", "tariff", "Protected/Non-Protected", "san_load", "status", "Current Bill", "Total FPA", "adjustments", "grand_total",
        "payment", "payment_date", "Total Payment", "Total Units",
        "ntn_no", "bill_month", "name_address", "due_date",
    ]
    ordered = [name for name in first if name in headers]
    meter_order = [
        "Meter Type",
        *(f"kwh_meter_{i}" for i in range(1, 5)),
        *(f"kwh_present_{i}" for i in range(1, 5)),
        *(f"kwh_mf_{i}" for i in range(1, 5)),
        *(f"kwh_units_{i}" for i in range(1, 5)),
        *(f"kvarh_present_{i}" for i in range(1, 5)),
        *(f"kvarh_units_{i}" for i in range(1, 5)),
        *(f"mdi_read_{i}" for i in range(1, 5)),
        *(f"mdi_mf_{i}" for i in range(1, 5)),
        *(f"mdi_units_{i}" for i in range(1, 5)),
        *(f"prv_cum_mdi_{i}" for i in range(1, 5)),
        *(f"cum_mdi_{i}" for i in range(1, 5)),
        "Diff_cum_mdi1",
        "diff_cum_mdi2",
        "Actual MDI",
        "Actual_Fix_Charges",
        "In_Bill_Fix_Charges (Estimated)",
    ]
    ordered += [name for name in meter_order if name in headers and name not in ordered]
    history_order = []
    for i in range(1, 13):
        history_order.extend([
            f"Month_{i}",
            f"Status_{i}",
            f"Units_{i}",
            f"Bill (Rs.)_{i}",
            f"Payment (Rs.)_{i}",
        ])
    ordered += [name for name in history_order if name in headers and name not in ordered]
    ordered += [name for name in headers if name not in ordered and name not in {"feeder", "Remarks", "scrape_status"}]
    if "feeder" in headers:
        ordered.append("feeder")
    if "Remarks" in headers or "scrape_status" in headers:
        ordered.append("Remarks")
    return ordered


def text_lines(element):
    """Parse distinct value lines from a table cell, separating (O)/(P) and space-separated units."""
    for br in element.find_all("br"):
        br.replace_with("\n")
    text = element.get_text("\n")
    t = re.sub(r"(?<=\S)\s+(\([A-Za-z0-9]+\))", r"\n\1", text)
    raw_lines = [clean(x) for x in t.splitlines() if clean(x)]
    result = []
    for line in raw_lines:
        nums = re.findall(r"[-+]?(?:\d+(?:\.\d+)?|\.\d+)", line)
        if len(nums) > 1 and not re.search(r"[A-Za-z]", line):
            result.extend(nums)
        else:
            result.append(line)
    return result


OFFICE_NS = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
TABLE_NS = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
TEXT_NS = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
SUPPORTED_EXTENSIONS = {".xlsx", ".ods"}


def _extract_ods_cell_value(cell):
    """Extract cell value from an ODF <table:table-cell> element."""
    val_type = cell.get(f"{{{OFFICE_NS}}}value-type")
    p_els = cell.findall(f"{{{TEXT_NS}}}p")
    text_val = "\n".join("".join(p.itertext()) for p in p_els).strip() if p_els else ""

    if val_type in {"float", "currency", "percentage"}:
        if re.fullmatch(r"0\d+", text_val):
            return text_val
        raw_val = cell.get(f"{{{OFFICE_NS}}}value")
        if raw_val is not None:
            try:
                dec = Decimal(raw_val)
                return int(dec) if dec == dec.to_integral_value() else float(dec)
            except Exception:
                pass
        return text_val if text_val else None

    if val_type == "boolean":
        bool_val = cell.get(f"{{{OFFICE_NS}}}boolean-value")
        if bool_val is not None:
            return bool_val.lower() == "true"
        return text_val if text_val else None

    if val_type in {"date", "time"}:
        if text_val:
            return text_val
        return cell.get(f"{{{OFFICE_NS}}}date-value") or cell.get(f"{{{OFFICE_NS}}}time-value")

    return text_val if text_val != "" else None


def _iter_ods_rows(container):
    """Yield <table:table-row> elements in document order without descending into nested tables."""
    row_tag = f"{{{TABLE_NS}}}table-row"
    row_group_tags = {
        f"{{{TABLE_NS}}}table-row-group",
        f"{{{TABLE_NS}}}table-header-rows",
        f"{{{TABLE_NS}}}table-rows",
    }
    for child in container:
        if child.tag == row_tag:
            yield child
        elif child.tag in row_group_tags:
            yield from _iter_ods_rows(child)


class _ODSCell:
    def __init__(self, value=None):
        self.value = value


class ODSSheet:
    """Provides openpyxl-compatible cell(row, col) access for ODS spreadsheets."""

    def __init__(self, table_el):
        self._table_el = table_el
        self._cells: dict[tuple[int, int], _ODSCell] = {}
        self.max_row = 0
        self.max_column = 0
        self._load()

    def _load(self):
        r_idx = 1
        for row_el in _iter_ods_rows(self._table_el):
            c_idx = 1
            for cell_el in row_el:
                if cell_el.tag not in (f"{{{TABLE_NS}}}table-cell", f"{{{TABLE_NS}}}covered-table-cell"):
                    continue
                rep_count = int(cell_el.get(f"{{{TABLE_NS}}}number-columns-repeated", "1"))
                val = _extract_ods_cell_value(cell_el)
                for _ in range(rep_count):
                    if val is not None and val != "":
                        self._cells[(r_idx, c_idx)] = _ODSCell(val)
                        if r_idx > self.max_row:
                            self.max_row = r_idx
                        if c_idx > self.max_column:
                            self.max_column = c_idx
                    c_idx += 1
            r_idx += 1
        if self.max_row == 0:
            self.max_row = max(r_idx - 1, 1)
        if self.max_column == 0:
            self.max_column = 1

    def cell(self, row, column):
        key = (row, column)
        if key not in self._cells:
            self._cells[key] = _ODSCell(None)
            if row > self.max_row:
                self.max_row = row
            if column > self.max_column:
                self.max_column = column
        return self._cells[key]

    def delete_cols(self, start, count=1):
        removals = [k for k in self._cells if start <= k[1] < start + count]
        for k in removals:
            del self._cells[k]
        shifted = {}
        for (r, c), obj in list(self._cells.items()):
            if c >= start + count:
                shifted[(r, c - count)] = obj
                del self._cells[(r, c)]
        self._cells.update(shifted)
        self.max_column = max([k[1] for k in self._cells] or [1])

    def delete_rows(self, start, count=1):
        removals = [k for k in self._cells if start <= k[0] < start + count]
        for k in removals:
            del self._cells[k]
        shifted = {}
        for (r, c), obj in list(self._cells.items()):
            if r >= start + count:
                shifted[(r - count, c)] = obj
                del self._cells[(r, c)]
        self._cells.update(shifted)
        self.max_row = max([k[0] for k in self._cells] or [1])


class ODSWorkbook:
    """Wrapper that reads and writes .ods files via zipfile and ElementTree."""

    def __init__(self, path):
        self.path = Path(path)
        with zipfile.ZipFile(self.path, "r") as zf:
            content_bytes = zf.read("content.xml")
        self._root = etree.fromstring(content_bytes)
        body = self._root.find(f".//{{{OFFICE_NS}}}body")
        spreadsheet = body.find(f"{{{OFFICE_NS}}}spreadsheet") if body is not None else None
        if spreadsheet is None:
            raise RuntimeError("Corrupt ODS: <office:spreadsheet> element not found.")
        self._table = spreadsheet.find(f"{{{TABLE_NS}}}table")
        if self._table is None:
            raise RuntimeError("Corrupt ODS: <table:table> element not found.")
        self.active = ODSSheet(self._table)

    def save(self, path=None):
        out_path = Path(path) if path else self.path
        sheet = self.active
        row_tag = f"{{{TABLE_NS}}}table-row"
        cell_tag = f"{{{TABLE_NS}}}table-cell"
        p_tag = f"{{{TEXT_NS}}}p"

        row_containers = {
            f"{{{TABLE_NS}}}table-row",
            f"{{{TABLE_NS}}}table-row-group",
            f"{{{TABLE_NS}}}table-header-rows",
            f"{{{TABLE_NS}}}table-rows",
        }
        col_containers = {
            f"{{{TABLE_NS}}}table-column",
            f"{{{TABLE_NS}}}table-column-group",
            f"{{{TABLE_NS}}}table-header-columns",
            f"{{{TABLE_NS}}}table-columns",
        }

        insert_idx = None
        for idx, child in enumerate(list(self._table)):
            if child.tag in row_containers:
                if insert_idx is None:
                    insert_idx = idx
                self._table.remove(child)

        if insert_idx is None:
            insert_idx = 0
            for idx, child in enumerate(self._table):
                if child.tag in col_containers:
                    insert_idx = idx + 1

        col_tag = f"{{{TABLE_NS}}}table-column"
        col_rep_attr = f"{{{TABLE_NS}}}number-columns-repeated"
        col_els = list(self._table.iter(col_tag))
        if col_els:
            for c in col_els:
                if col_rep_attr in c.attrib:
                    del c.attrib[col_rep_attr]
            while len(col_els) < sheet.max_column:
                new_col = etree.Element(col_tag)
                col_els[-1].addnext(new_col)
                col_els.append(new_col)
            for c in col_els[sheet.max_column:]:
                parent = c.getparent()
                if parent is not None:
                    parent.remove(c)

        for r in range(1, sheet.max_row + 1):
            row_el = etree.Element(row_tag)
            for c in range(1, sheet.max_column + 1):
                cell_el = etree.SubElement(row_el, cell_tag)
                val = sheet.cell(r, c).value
                if val is not None and val != "":
                    if isinstance(val, (int, float, Decimal)):
                        cell_el.set(f"{{{OFFICE_NS}}}value-type", "float")
                        cell_el.set(f"{{{OFFICE_NS}}}value", str(val))
                        p_el = etree.SubElement(cell_el, p_tag)
                        p_el.text = str(val)
                    elif isinstance(val, bool):
                        cell_el.set(f"{{{OFFICE_NS}}}value-type", "boolean")
                        cell_el.set(f"{{{OFFICE_NS}}}boolean-value", "true" if val else "false")
                        p_el = etree.SubElement(cell_el, p_tag)
                        p_el.text = str(val).lower()
                    else:
                        cell_el.set(f"{{{OFFICE_NS}}}value-type", "string")
                        for line in str(val).split("\n"):
                            p_el = etree.SubElement(cell_el, p_tag)
                            p_el.text = line
            self._table.insert(insert_idx, row_el)
            insert_idx += 1

        new_content = etree.tostring(self._root, xml_declaration=True, encoding="UTF-8")

        buf = io.BytesIO()
        with zipfile.ZipFile(self.path, "r") as zin:
            with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == "content.xml":
                        zout.writestr(item, new_content)
                    else:
                        zout.writestr(item, zin.read(item.filename))
        out_path.write_bytes(buf.getvalue())


def load_spreadsheet(path):
    suffix = Path(path).suffix.lower()
    if suffix == ".xlsx":
        return load_workbook(path)
    if suffix == ".ods":
        return ODSWorkbook(path)
    raise ValueError(f"Unsupported spreadsheet format: {suffix}")


class PITCBillScraper:
    def __init__(self, delay=2.0, timeout=90, history_mode=None, url=None):
        self.delay = delay
        self.timeout = timeout
        self.history_mode = history_mode
        self.url = (url or URL).strip().rstrip('/')

    @staticmethod
    def put(output, name, value):
        value = clean(value)
        if not value or is_dropped(name):
            return
        # Never accept a heading text as a data value
        norm_val = re.sub(r"[^a-z0-9]+", " ", value.lower()).strip()
        if norm_val in BILL_HEADING_KEYWORDS or clean(value).upper() == clean(name).upper():
            return
        key = canonical_label(name)

        if key == "adjustments":
            numeric_adjustment = re.search(r"[-+]?\d[\d,]*(?:\.\d+)?", value)
            if not numeric_adjustment:
                return
            value = numeric_adjustment.group(0)

        if re.search(r"(?:meter|present|mdi|kwh|kvarh|reading)", key, re.I):
            value = re.sub(r"^\s*[IE]\s+", "", value, flags=re.I).strip()

        if "units" in key:
            value = re.sub(r"\(\s*[OP]\s*\)\s*", "", value, flags=re.I).strip()

        # If a units field receives multiple space-separated numbers (e.g. '1673 610'),
        # automatically distribute them into index and index+1 so they are never merged into one cell.
        units_match = re.match(r"^(kwh|kvarh|mdi)_units_?(\d+)?$", key)
        if units_match:
            prefix = units_match.group(1)
            start_idx = int(units_match.group(2)) if units_match.group(2) else 1
            nums = [clean(n) for n in re.findall(r"[-+]?(?:\d+(?:\.\d+)?|\.\d+)", value)]
            if len(nums) > 1:
                for offset, num in enumerate(nums):
                    target_key = f"{prefix}_units_{start_idx + offset}"
                    existing = output.get(target_key)
                    if existing is None or (str(existing).strip() in {"0", "0.0", "0.00"} and str(num).strip() not in {"0", "0.0", "0.00"}):
                        output[target_key] = num
                return

        if value:
            existing = output.get(key)
            if existing is None or (str(existing).strip() in {"0", "0.0", "0.00"} and str(value).strip() not in {"0", "0.0", "0.00"}):
                output[key] = value

    @staticmethod
    def put_grid_col_cells(output, soup):
        """Read 1-to-1 label-row and val-space pairs inside .grid-col-cell."""
        for cell in soup.select(".grid-col-cell"):
            label_rows = cell.select(".label-row")
            val_spaces = cell.select(".val-space")
            for lr, vs in zip(label_rows, val_spaces):
                en_lbl = lr.select_one(".en-lbl")
                lbl_text = clean(en_lbl.get_text(" ")) if en_lbl else clean(lr.get_text(" "))
                val_text = clean(" | ".join(text_lines(vs)))
                canon = canonical_label(lbl_text)
                if val_text:
                    PITCBillScraper.put(output, lbl_text, val_text)
                else:
                    if canon not in output:
                        output[canon] = ""

    @staticmethod
    def put_general_pairs(output, soup):
        """Read label/value pairs from both MDI and ordinary bill layouts."""
        known = set(GENERAL_LABELS)
        for row in soup.select("tr"):
            cells = [clean(x.get_text(" ")) for x in row.select("th, td") if clean(x.get_text(" "))]
            if len(cells) >= 2:
                label = cells[0].lower().rstrip(":")
                if label in known:
                    PITCBillScraper.put(output, cells[0], cells[-1])
        for node in soup.select(".charges-bd-row, .right-section-cell"):
            lbl_el = node.select_one(".right-panel-en, .charges-bd-en, .charges-bd-label")
            val_el = node.select_one(".right-main-val, .charges-bd-val")
            if lbl_el and val_el:
                lbl = clean(lbl_el.get_text(" ")).lower().rstrip(":")
                lbl = re.sub(r"^[^a-z0-9]+", "", lbl).strip()
                val = clean(val_el.get_text(" "))
                if lbl in known:
                    PITCBillScraper.put(output, lbl, val)
            else:
                bits = [clean(x) for x in node.stripped_strings if clean(x)]
                if len(bits) >= 2 and bits[0].lower().rstrip(":") in known:
                    PITCBillScraper.put(output, bits[0], bits[-1])

    @staticmethod
    def put_general_text_pairs(output, soup):
        """Fallback for general batches whose labels are not table cells."""
        main = soup.select_one(".gbn-app-shell") or soup.body or soup
        if not main:
            return
        labels = set(GENERAL_LABELS)
        lines = [clean(x) for x in main.stripped_strings if clean(x)]
        for index, line in enumerate(lines):
            label = re.sub(r"[^a-z0-9]+", " ", line.lower()).strip()
            if label not in labels:
                continue
            key = GENERAL_LABELS[label]
            if is_dropped(key):
                continue
            if key in output:
                existing = str(output[key])
                if key == "san_load" and re.search(r"[A-Za-z]-?\d", existing):
                    del output[key]
                elif key == "tariff" and not re.search(r"[A-Za-z]-?\d", existing):
                    del output[key]
                else:
                    continue
            for candidate in lines[index + 1:index + 4]:
                candidate_key = re.sub(r"[^a-z0-9]+", " ", candidate.lower()).strip()
                if candidate_key in BILL_HEADING_KEYWORDS or candidate_key in labels or not candidate:
                    break
                if any(ord(ch) > 127 for ch in candidate) and not re.search(r"\d", candidate):
                    continue
                if key == "san_load" and not re.search(r"\d", candidate):
                    continue
                if key == "tariff" and not re.search(r"[A-Za-z]-?\d", candidate):
                    continue
                PITCBillScraper.put(output, key, candidate)
                break

    @staticmethod
    def extract_history(soup, mode):
        """Extract billing and payment history from history grid boxes.

        mode == '1yr': Box 0 (skip 1st month) + Box 1 (all months) -> 11 historical months
        mode == '6mo': Box 1 (.history-center-divider, skip 1st month) -> 5 historical months
        """
        wrapper = soup.select_one(".history-data-wrapper")
        grids = wrapper.select(".history-block-grid") if wrapper else soup.select(".history-block-grid")
        if not grids:
            return {}

        target_rows = []
        if mode == "1yr":
            box0_rows = grids[0].select(".history-row")
            if len(box0_rows) > 1:
                target_rows.extend(box0_rows[1:])
            elif box0_rows:
                target_rows.extend(box0_rows)

            if len(grids) > 1:
                target_rows.extend(grids[1].select(".history-row"))
        elif mode == "6mo":
            box1 = None
            for g in grids:
                classes = g.get("class", [])
                if "history-center-divider" in classes:
                    box1 = g
                    break
            if not box1 and len(grids) > 1:
                box1 = grids[1]
            elif not box1 and grids:
                box1 = grids[0]

            if box1:
                box1_rows = box1.select(".history-row")
                if len(box1_rows) > 1:
                    target_rows.extend(box1_rows[1:])
                elif box1_rows:
                    target_rows.extend(box1_rows)

        out = {}
        for idx, row in enumerate(target_rows, 1):
            cells = [clean(c.get_text(" ", strip=True)) for c in row.select(".history-cell")]
            out[f"Month_{idx}"] = cells[0] if len(cells) > 0 else ""
            out[f"Status_{idx}"] = cells[1] if len(cells) > 1 else ""
            out[f"Units_{idx}"] = cells[2] if len(cells) > 2 else ""
            out[f"Bill (Rs.)_{idx}"] = cells[3] if len(cells) > 3 else ""
            out[f"Payment (Rs.)_{idx}"] = cells[4] if len(cells) > 4 else ""
        return out

    @staticmethod
    def put_meter_tables(output, soup):
        """Map both MDI and ordinary meter tables to stable reading headings."""
        header_aliases = {
            "meter": "meter", "meter no": "meter", "meter no.": "meter", "meter number": "meter",
            "present": "present", "present reading": "present", "present readings": "present",
            "mf": "mf", "units": "units", "unit": "units", "type": None,
            "previous": None, "previous reading": None, "previous readings": None,
        }
        for table in soup.select("table"):
            headers = [clean(x.get_text(" ")).lower() for x in table.select("thead th")]
            if not headers:
                headers = [clean(x.get_text(" ")).lower() for x in table.select("tr:first-child th, tr:first-child td")]
            mapped = [header_aliases.get(re.sub(r"\s+", " ", h).strip()) for h in headers]
            if not any(x in {"meter", "present", "mf", "units"} for x in mapped):
                continue
            rows = table.select("tbody tr") or table.select("tr")[1:]
            carry = None
            for row in rows:
                cells = row.select("th, td")
                if not cells:
                    continue
                if headers and len(cells) == len(headers) - 1 and carry is not None:
                    cells = [carry] + cells
                if cells and headers and headers[0] in {"meter", "meter no", "meter no."}:
                    carry = cells[0]
                row_text = " ".join(clean(x.get_text(" ")) for x in cells[:2]).lower()
                prefix = "kvarh" if "kvarh" in row_text or "kvar h" in row_text else "mdi" if re.search(r"\bmdi\b", row_text) else "kwh"
                for col_no, cell in enumerate(cells):
                    kind = mapped[col_no] if col_no < len(mapped) else None
                    if kind is None:
                        continue
                    for value_no, value in enumerate(text_lines(cell), 1):
                        self_value = re.sub(r"\b[IE]\s+(?=\d)", "", value, flags=re.I).strip()
                        PITCBillScraper.put(output, f"{prefix}_{kind}_{value_no}", self_value)

    @staticmethod
    def put_meter_text_sections(output, soup):
        """Fallback for general bills that render meter data as div/text blocks."""
        main = soup.select_one(".gbn-app-shell") or soup.body or soup
        lines = [clean(x) for x in main.stripped_strings if clean(x)]
        lowered = [re.sub(r"[^a-z0-9]+", " ", x.lower()).strip() for x in lines]
        try:
            start = next(i for i, value in enumerate(lowered) if value in {"meter info", "meter information"})
        except StopIteration:
            return
        end = next((i for i in range(start + 1, len(lines)) if lowered[i] in {"bill charges breakdown", "bill history"}), len(lines))
        section_lines = lines[start:end]
        section_keys = [re.sub(r"[^a-z0-9]+", " ", x.lower()).strip() for x in section_lines]
        if not any(value in {"meter no", "meter number"} for value in section_keys):
            return
        labels = {
            "meter no": "meter", "meter number": "meter", "meter": "meter",
            "present reading": "present", "present readings": "present",
            "present": "present", "mf": "mf", "units": "units", "unit": "units",
        }
        for index, label in enumerate(section_keys):
            kind = labels.get(label)
            if not kind:
                continue
            values = []
            for candidate in section_lines[index + 1:]:
                candidate_key = re.sub(r"[^a-z0-9]+", " ", candidate.lower()).strip()
                if candidate_key in labels or candidate_key in {"previous", "previous reading", "type", "meter info"}:
                    break
                if any(ord(ch) > 127 for ch in candidate) and not re.search(r"\d", candidate):
                    continue
                if kind == "meter" and not re.search(r"\d", candidate):
                    continue
                if kind != "meter" and not re.search(r"[-+]?\d", candidate):
                    continue
                values.append(re.sub(r"\b[IE]\s+(?=\d)", "", candidate, flags=re.I).strip())
            for value_no, value in enumerate(values, 1):
                PITCBillScraper.put(output, f"kwh_{kind}_{value_no}", value)

    @staticmethod
    def put_cum_mdi(output, soup):
        """Read previous and present cumulative MDI values."""
        # 1. Check explicit MDI meter footer cells if present
        for cell in soup.select(".mdi-meter-footer-cell"):
            lbl_el = cell.select_one(".en-lbl")
            val_el = cell.select_one(".mdi-meter-footer-val")
            if lbl_el and val_el:
                lbl = clean(lbl_el.get_text(" ")).upper()
                vals = [clean(x) for x in text_lines(val_el) if re.fullmatch(r"[-+]?(?:\d+(?:\.\d+)?|\.\d+)", clean(x))]
                if "PRV" in lbl and "CUM" in lbl:
                    for idx, val in enumerate(vals[:4], 1):
                        output[f"prv_cum_mdi_{idx}"] = val
                elif "PRS" in lbl and "CUM" in lbl:
                    for idx, val in enumerate(vals[:4], 1):
                        output[f"cum_mdi_{idx}"] = val

        # 2. Supplementary scan through stripped lines fallback
        main = soup.select_one(".gbn-app-shell") or soup.body or soup
        if not main:
            return
        lines = [clean(x) for x in main.stripped_strings if clean(x)]
        keys = [re.sub(r"[^a-z0-9]+", " ", x.lower()).strip() for x in lines]

        # PRV CUMM MDI
        prv_start = next((i for i, value in enumerate(keys) if value in {"prv cumm mdi", "previous cumm mdi", "previous cumulative mdi", "prv cum mdi"}), None)
        if prv_start is not None:
            values = []
            for candidate in lines[prv_start + 1:prv_start + 7]:
                if re.fullmatch(r"[-+]?(?:\d+(?:\.\d+)?|\.\d+)", candidate):
                    values.append(candidate)
                elif values:
                    break
            for value_no, value in enumerate(values[:4], 1):
                k = f"prv_cum_mdi_{value_no}"
                if k not in output:
                    output[k] = value

        # PRS CUMM MDI
        prs_start = next((i for i, value in enumerate(keys) if value in {"prs cumm mdi", "present cumm mdi", "present cumulative mdi", "prs cum mdi"}), None)
        if prs_start is not None:
            values = []
            for candidate in lines[prs_start + 1:prs_start + 7]:
                if re.fullmatch(r"[-+]?(?:\d+(?:\.\d+)?|\.\d+)", candidate):
                    values.append(candidate)
                elif values:
                    break
            for value_no, value in enumerate(values[:4], 1):
                k = f"cum_mdi_{value_no}"
                if k not in output:
                    output[k] = value

    @staticmethod
    def put_charges_breakdown(output, soup):
        """Extract Current Bill and Total FPA from charges breakdown."""
        # 1. Look for .charges-bd-row containers
        for row in soup.select(".charges-bd-row, tr"):
            row_text = clean(row.get_text(" ")).upper()

            # Current Bill (class=charges-bd-val)
            if "CURRENT BILL" in row_text:
                val_el = row.select_one(".charges-bd-val")
                if val_el:
                    v = clean(val_el.get_text(" "))
                    if v and re.search(r"\d", v):
                        output["Current Bill"] = v
                else:
                    bits = [clean(x) for x in row.stripped_strings if clean(x)]
                    if len(bits) >= 2 and re.search(r"\d", bits[-1]):
                        output["Current Bill"] = bits[-1]

            # Total FPA (class=charges-bd-val charges-bd-val--strong)
            if any(term in row_text for term in ("TOTAL F.P.A", "TOTAL FPA", "TOTAL F P A", "TOTAL FUEL PRICE")):
                val_el = (
                    row.select_one(".charges-bd-val.charges-bd-val--strong")
                    or row.select_one(".charges-bd-val--strong")
                    or row.select_one(".charges-bd-val")
                )
                if val_el:
                    v = clean(val_el.get_text(" "))
                    if v and re.search(r"\d", v):
                        output["Total FPA"] = v
                else:
                    bits = [clean(x) for x in row.stripped_strings if clean(x)]
                    if len(bits) >= 2 and re.search(r"\d", bits[-1]):
                        output["Total FPA"] = bits[-1]

        # 2. Target direct CSS classes if not yet populated
        if "Total FPA" not in output:
            for el in soup.select(".charges-bd-val--strong, .charges-bd-val.charges-bd-val--strong"):
                parent = el.find_parent(class_="charges-bd-row") or el.parent
                p_text = clean(parent.get_text(" ")).upper() if parent else ""
                if any(term in p_text for term in ("FPA", "F.P.A", "FUEL")):
                    v = clean(el.get_text(" "))
                    if v and re.search(r"\d", v):
                        output["Total FPA"] = v
                        break

        # 3. Fallback scan through stripped lines
        main = soup.select_one(".gbn-app-shell") or soup.body or soup
        if main and ("Current Bill" not in output or "Total FPA" not in output):
            lines = [clean(x) for x in main.stripped_strings if clean(x)]
            for idx, line in enumerate(lines):
                norm_line = re.sub(r"[^A-Z0-9]+", " ", line.upper()).strip()
                if "Current Bill" not in output and norm_line in {"CURRENT BILL", "CURRENT BILL AMOUNT"}:
                    for cand in lines[idx + 1:idx + 4]:
                        if re.search(r"[-+]?\d[\d,]*(?:\.\d+)?", cand):
                            output["Current Bill"] = cand
                            break
                if "Total FPA" not in output and ("TOTAL F P A" in norm_line or "TOTAL FPA" in norm_line):
                    for cand in lines[idx + 1:idx + 4]:
                        if re.search(r"[-+]?\d[\d,]*(?:\.\d+)?", cand):
                            output["Total FPA"] = cand
                            break

    def _fetch_url_with_retry(self, session, method, url, data=None, stop_event=None, log_callback=None):
        was_disconnected = False
        while True:
            if stop_event and stop_event.is_set():
                raise InterruptedError("Operation stopped by user")
            try:
                if method.lower() == "post":
                    resp = session.post(url, data=data, timeout=self.timeout)
                else:
                    resp = session.get(url, timeout=self.timeout)
                resp.raise_for_status()
                if was_disconnected and log_callback:
                    log_callback("Internet connection restored. Resuming bill processing...")
                return resp
            except (requests.exceptions.ConnectionError, requests.exceptions.Timeout, ConnectionResetError, OSError):
                was_disconnected = True
                if log_callback:
                    log_callback("Internet disconnected. Waiting for network connection to restore...")
                for _ in range(30):
                    if stop_event and stop_event.is_set():
                        raise InterruptedError("Operation stopped by user")
                    time.sleep(0.1)
            except requests.exceptions.HTTPError as exc:
                if exc.response is not None and 500 <= exc.response.status_code <= 504:
                    was_disconnected = True
                    if log_callback:
                        log_callback(f"PITC server busy (HTTP {exc.response.status_code}). Retrying in a moment...")
                    for _ in range(20):
                        if stop_event and stop_event.is_set():
                            raise InterruptedError("Operation stopped by user")
                        time.sleep(0.1)
                    continue
                raise

    def search(self, account, stop_event=None, log_callback=None):
        session = requests.Session()
        session.headers.update({"User-Agent": "Mozilla/5.0 (compatible; PITC-Bill-Exporter/1.0)"})
        target_url = getattr(self, "url", URL)
        page = self._fetch_url_with_retry(session, "get", target_url, stop_event=stop_event, log_callback=log_callback)
        try:
            import lxml
            parser = "lxml"
        except ImportError:
            parser = "html.parser"
        soup = BeautifulSoup(page.text, parser)
        form = soup.select_one("form#SubmitForm") or soup.find("form")
        if not form:
            raise RuntimeError("PITC search form was not found")
        payload = {i["name"]: i.get("value", "") for i in form.select("input[type=hidden][name]")}
        payload.update({"rbSearchByList": "refno", "searchTextBox": account, "ruCodeTextBox": "", "btnSearch": "Search"})
        result = self._fetch_url_with_retry(session, "post", target_url, data=payload, stop_event=stop_event, log_callback=log_callback)
        if "CONSUMER DETAIL" not in result.text.upper() and "CONSUMER DETAILS" not in result.text.upper():
            raise RuntimeError(clean(BeautifulSoup(result.text, parser).get_text(" "))[:300] or "No bill returned")
        return BeautifulSoup(result.text, parser)

    def scrape_account(self, raw_account, existing_status=None, stop_event=None, log_callback=None, history_mode=None):
        if history_mode is None:
            history_mode = self.history_mode
        raw = str(raw_account).strip()
        if re.fullmatch(r"[0-9.]+e[+\-]?[0-9]+", raw, re.I):
            raw = format(Decimal(raw), "f")
        account = re.sub(r"\D", "", raw)
        clean_ac = re.sub(r"\D", "", raw)
        bno = clean_ac[:2] if len(clean_ac) >= 2 else ""
        sdiv = clean_ac[2:7] if len(clean_ac) >= 7 else (clean_ac[2:] if len(clean_ac) > 2 else "")

        if not re.fullmatch(r"\d{10,14}", account):
            return {
                "_account": account,
                "BNo": bno,
                "SDiv": sdiv,
                "Protected/Non-Protected": "",
                "Total Payment": "",
                "Total Units": "",
                "Meter Type": "",
                "Remarks": "Bill Not Found",
                "scrape_status": "Bill Not Found",
            }
        try:
            soup = self.search(account, stop_event=stop_event, log_callback=log_callback)
            out = {
                "_account": account,
                "BNo": bno,
                "SDiv": sdiv,
                "Remarks": "Fetched",
                "scrape_status": "Fetched",
            }

            # 1. Structured grid cells (Status, Tariff, San Load, Name, Feeder, etc.)
            self.put_grid_col_cells(out, soup)

            # 2. Key-value pairs
            self.put_general_pairs(out, soup)

            # 3. Fallback text pairs
            self.put_general_text_pairs(out, soup)

            self.put_meter_tables(out, soup)
            self.put_meter_text_sections(out, soup)
            self.put_cum_mdi(out, soup)
            self.put_charges_breakdown(out, soup)

            # Final validation: ensure units fields don't have multiple numbers in a single key
            for prefix in ("kwh", "kvarh", "mdi"):
                for idx in range(1, 5):
                    k = f"{prefix}_units_{idx}"
                    if k in out and out[k]:
                        val_str = str(out[k])
                        nums = [clean(n) for n in re.findall(r"[-+]?(?:\d+(?:\.\d+)?|\.\d+)", val_str)]
                        if len(nums) > 1:
                            out[k] = nums[0]
                            next_k = f"{prefix}_units_{idx + 1}"
                            if next_k not in out or not out[next_k]:
                                out[next_k] = nums[1]

            # Extract Meter Type from kwh_meter_1 if present
            m_type, m_no = extract_meter_type_and_no(out.get("kwh_meter_1"))
            if m_type:
                out["Meter Type"] = m_type
                out["kwh_meter_1"] = m_no
            elif "Meter Type" not in out or out["Meter Type"] is None:
                out["Meter Type"] = ""

            # Billing/Payment History extraction if requested
            if history_mode in ("1yr", "6mo"):
                hist_data = self.extract_history(soup, history_mode)
                out.update(hist_data)

            # Sanitize stray 'TARIFF CATEGORY' from status
            if str(out.get("status", "")).strip().upper() == "TARIFF CATEGORY":
                out["status"] = ""

            # Protected/Non-Protected calculation:
            # If consumer's tariff is A-1A(01) and he has not consumed above 200 units in last 6 months
            # (current month + 5 months from history) then put "Protected", otherwise "Non-Protected".
            h_6mo_data = self.extract_history(soup, "6mo")
            curr_u = parse_num(out.get("kwh_units_1"))
            if curr_u is None:
                curr_u = parse_num(out.get("kwh_units"))
            if curr_u is None:
                curr_u = 0.0
            h_units = [parse_num(h_6mo_data.get(f"Units_{i}")) for i in range(1, 6)]
            out["Protected/Non-Protected"] = compute_protected_status(out.get("tariff"), curr_u, h_units)

            # Total Payment calculation: sum of all payments already fetched in excel
            out["Total Payment"] = compute_total_payment(out)

            # Total Units calculation: sum of all units consumed fetched in excel
            out["Total Units"] = compute_total_units(out)

            def format_num(num, decimals=4):
                if num is None:
                    return None
                val = round(num, decimals)
                return int(val) if val.is_integer() else val

            # Check status:
            # 1. If P.Disc or T.Disc: skip subtraction, Actual MDI, and actual fixed charges.
            # 2. If Replaced: subtract (cum_mdi_2 - prv_cum_mdi_2) and (cum_mdi_4 - prv_cum_mdi_4).
            # 3. Normal / Other: subtract (cum_mdi_1 - prv_cum_mdi_1) and (cum_mdi_2 - prv_cum_mdi_2).
            status_val = out.get("status") or existing_status

            if is_disc_status(status_val):
                out["Diff_cum_mdi1"] = None
                out["diff_cum_mdi2"] = None
                out["Actual MDI"] = None
                out["Actual_Fix_Charges"] = None
            else:
                if is_replaced_status(status_val):
                    # Case Replaced: cumm mdi 2 - prv cumm mdi 2, and cumm mdi 4 - prv cumm mdi 4
                    prv2 = parse_num(out.get("prv_cum_mdi_2"))
                    cum2 = parse_num(out.get("cum_mdi_2"))
                    diff1 = format_num(abs(cum2 - prv2), 4) if (prv2 is not None and cum2 is not None) else None
                    if diff1 is not None:
                        out["Diff_cum_mdi1"] = diff1

                    prv4 = parse_num(out.get("prv_cum_mdi_4"))
                    cum4 = parse_num(out.get("cum_mdi_4"))
                    diff2 = format_num(abs(cum4 - prv4), 4) if (prv4 is not None and cum4 is not None) else None
                    if diff2 is not None:
                        out["diff_cum_mdi2"] = diff2
                else:
                    # Normal status: cumm mdi 1 - prv cumm mdi 1, and cumm mdi 2 - prv cumm mdi 2
                    prv1 = parse_num(out.get("prv_cum_mdi_1"))
                    cum1 = parse_num(out.get("cum_mdi_1"))
                    diff1 = format_num(abs(cum1 - prv1), 4) if (prv1 is not None and cum1 is not None) else None
                    if diff1 is not None:
                        out["Diff_cum_mdi1"] = diff1

                    prv2 = parse_num(out.get("prv_cum_mdi_2"))
                    cum2 = parse_num(out.get("cum_mdi_2"))
                    diff2 = format_num(abs(cum2 - prv2), 4) if (prv2 is not None and cum2 is not None) else None
                    if diff2 is not None:
                        out["diff_cum_mdi2"] = diff2

                # Actual MDI: max(Diff_cum_mdi1, diff_cum_mdi2) * mdi_mf_1
                diff_candidates = []
                for d in (out.get("Diff_cum_mdi1"), out.get("diff_cum_mdi2")):
                    if d is not None:
                        diff_candidates.append(float(d))

                if diff_candidates:
                    max_diff = max(diff_candidates)
                    mf = parse_num(out.get("mdi_mf_1"))
                    if mf is None or mf == 0:
                        mf = 1.0
                    actual_mdi = format_num(max_diff * mf, 4)
                    out["Actual MDI"] = actual_mdi

                    # Actual_Fix_Charges = Actual MDI x 1250
                    actual_fix = format_num(float(actual_mdi) * 1250, 2)
                    out["Actual_Fix_Charges"] = actual_fix

            # In_Bill_Fix_Charges (Estimated): mdi_units_1 multiplied by 1250
            mdi_u1 = parse_num(out.get("mdi_units_1") if "mdi_units_1" in out else out.get("mdi_units"))
            if mdi_u1 is not None:
                in_bill_fix = format_num(mdi_u1 * 1250, 2)
                out["In_Bill_Fix_Charges (Estimated)"] = in_bill_fix

            return out
        except InterruptedError:
            raise
        except Exception as exc:
            return {
                "_account": account,
                "BNo": bno,
                "SDiv": sdiv,
                "Protected/Non-Protected": "",
                "Total Payment": "",
                "Total Units": "",
                "Meter Type": "",
                "Remarks": "Bill Not Found",
                "scrape_status": "Bill Not Found",
            }

    def update_workbook(self, path, progress=None, stop_event=None, log_callback=None):
        if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            raise RuntimeError("Please select an .xlsx or .ods workbook. The selected workbook is updated in place.")
        workbook = load_spreadsheet(path)
        sheet = workbook.active

        first = str(sheet.cell(1, 1).value or "").strip().lower().rstrip(".")
        has_heading = first in {"ac", "ac no", "account", "account no"}
        data_start_row = 2 if has_heading else 1

        old_max_col = sheet.max_column
        old_max_row = sheet.max_row

        existing_col_names = {}
        if has_heading:
            for col in range(1, old_max_col + 1):
                raw = sheet.cell(1, col).value
                if raw:
                    norm = "AC No." if col == 1 else normalized_key(raw)
                    if norm.lower() in ("scrape_status", "scrape status"):
                        norm = "Remarks"
                    existing_col_names[col] = norm

        row_records = []
        for r in range(data_start_row, old_max_row + 1):
            raw_ac = sheet.cell(r, 1).value
            if raw_ac is None or not str(raw_ac).strip():
                continue
            record = {}
            for col in range(1, old_max_col + 1):
                val = sheet.cell(r, col).value
                col_name = existing_col_names.get(col)
                if col_name and val is not None:
                    if col_name == "status" and str(val).strip().upper() == "TARIFF CATEGORY":
                        val = ""
                    record[col_name] = val
            record["AC No."] = raw_ac
            row_records.append(record)

        total = len(row_records)
        all_keys = set()
        for col_name in existing_col_names.values():
            if not is_dropped(col_name) and col_name not in ("scrape_status", "scrape status"):
                all_keys.add(col_name)

        # Pre-populate BNo, SDiv, Meter Type, Protected/Non-Protected, Total Payment, and Total Units on all records
        for record in row_records:
            ac_clean = re.sub(r"\D", "", str(record.get("AC No.", "")))
            if not record.get("BNo"):
                record["BNo"] = ac_clean[:2] if len(ac_clean) >= 2 else ""
            if not record.get("SDiv") or (len(str(record.get("SDiv", "")).strip()) < 4 and len(ac_clean) >= 7):
                record["SDiv"] = ac_clean[2:7] if len(ac_clean) >= 7 else (ac_clean[2:] if len(ac_clean) > 2 else "")
            m_type, m_no = extract_meter_type_and_no(record.get("kwh_meter_1"))
            if m_type:
                record["Meter Type"] = m_type
                record["kwh_meter_1"] = m_no
            elif "Meter Type" not in record or record["Meter Type"] is None:
                record["Meter Type"] = ""
            clean_t = re.sub(r"\s+", "", str(record.get("tariff", ""))).upper()
            if clean_t != "A-1A(01)":
                record["Protected/Non-Protected"] = ""
            elif not record.get("Protected/Non-Protected"):
                curr_u = parse_num(record.get("kwh_units_1"))
                if curr_u is None:
                    curr_u = parse_num(record.get("kwh_units"))
                if curr_u is None:
                    curr_u = 0.0
                h_units = [parse_num(record.get(f"Units_{i}")) for i in range(1, 6)]
                record["Protected/Non-Protected"] = compute_protected_status(record.get("tariff"), curr_u, h_units)
            if "Total Payment" not in record or record["Total Payment"] is None:
                record["Total Payment"] = compute_total_payment(record)
            if "Total Units" not in record or record["Total Units"] is None:
                record["Total Units"] = compute_total_units(record)

        try:
            for done, record in enumerate(row_records, 1):
                if stop_event and stop_event.is_set():
                    break
                account = record.get("AC No.")
                existing_status = record.get("status")
                try:
                    result = self.scrape_account(
                        account,
                        existing_status=existing_status,
                        stop_event=stop_event,
                        log_callback=log_callback,
                        history_mode=self.history_mode,
                    )
                except InterruptedError:
                    break
                record.update(result)

                ac_clean = re.sub(r"\D", "", str(account))
                if not record.get("BNo"):
                    record["BNo"] = ac_clean[:2] if len(ac_clean) >= 2 else ""
                if not record.get("SDiv") or (len(str(record.get("SDiv", "")).strip()) < 4 and len(ac_clean) >= 7):
                    record["SDiv"] = ac_clean[2:7] if len(ac_clean) >= 7 else (ac_clean[2:] if len(ac_clean) > 2 else "")

                m_type, m_no = extract_meter_type_and_no(record.get("kwh_meter_1"))
                if m_type:
                    record["Meter Type"] = m_type
                    record["kwh_meter_1"] = m_no
                elif "Meter Type" not in record or record["Meter Type"] is None:
                    record["Meter Type"] = ""

                if str(record.get("status", "")).strip().upper() == "TARIFF CATEGORY":
                    record["status"] = ""

                effective_status = record.get("status") or existing_status
                if is_disc_status(effective_status):
                    record["Diff_cum_mdi1"] = None
                    record["diff_cum_mdi2"] = None
                    record["Actual MDI"] = None
                    record["Actual_Fix_Charges"] = None

                record["Total Payment"] = compute_total_payment(record)
                record["Total Units"] = compute_total_units(record)

                clean_t = re.sub(r"\s+", "", str(record.get("tariff", ""))).upper()
                if clean_t != "A-1A(01)":
                    record["Protected/Non-Protected"] = ""
                elif not record.get("Protected/Non-Protected"):
                    curr_u = parse_num(record.get("kwh_units_1"))
                    if curr_u is None:
                        curr_u = parse_num(record.get("kwh_units"))
                    if curr_u is None:
                        curr_u = 0.0
                    h_units = [parse_num(record.get(f"Units_{i}")) for i in range(1, 6)]
                    record["Protected/Non-Protected"] = compute_protected_status(record.get("tariff"), curr_u, h_units)

                for k in result:
                    if not k.startswith("_") and not is_dropped(k) and k not in ("scrape_status", "scrape status"):
                        all_keys.add(k)
                if progress:
                    progress(done, total, record)

                if stop_event and stop_event.is_set():
                    break

                slept = 0.0
                delay_step = 0.1
                while slept < self.delay:
                    if stop_event and stop_event.is_set():
                        break
                    time.sleep(delay_step)
                    slept += delay_step
        finally:
            all_keys.add("AC No.")
            all_keys.add("BNo")
            all_keys.add("SDiv")
            all_keys.add("tariff")
            all_keys.add("Protected/Non-Protected")
            all_keys.add("san_load")
            all_keys.add("status")
            all_keys.add("Current Bill")
            all_keys.add("Total FPA")
            all_keys.add("adjustments")
            all_keys.add("grand_total")
            all_keys.add("payment")
            all_keys.add("payment_date")
            all_keys.add("Total Payment")
            all_keys.add("Total Units")
            all_keys.add("Meter Type")
            all_keys.add("feeder")
            all_keys.add("Remarks")
            all_keys.discard("scrape_status")
            all_keys.discard("scrape status")

            if self.history_mode == "1yr":
                for i in range(1, 12):
                    all_keys.add(f"Month_{i}")
                    all_keys.add(f"Status_{i}")
                    all_keys.add(f"Units_{i}")
                    all_keys.add(f"Bill (Rs.)_{i}")
                    all_keys.add(f"Payment (Rs.)_{i}")
                for k in list(all_keys):
                    m = re.match(r"^(?:Month|Status|Units|Bill \(Rs\.\)|Payment \(Rs\.\))_(\d+)$", k)
                    if m and int(m.group(1)) > 11:
                        all_keys.discard(k)
            elif self.history_mode == "6mo":
                for i in range(1, 6):
                    all_keys.add(f"Month_{i}")
                    all_keys.add(f"Status_{i}")
                    all_keys.add(f"Units_{i}")
                    all_keys.add(f"Bill (Rs.)_{i}")
                    all_keys.add(f"Payment (Rs.)_{i}")
                for k in list(all_keys):
                    m = re.match(r"^(?:Month|Status|Units|Bill \(Rs\.\)|Payment \(Rs\.\))_(\d+)$", k)
                    if m and int(m.group(1)) > 5:
                        all_keys.discard(k)
            else:
                for k in list(all_keys):
                    if re.match(r"^(?:Month|Status|Units|Bill \(Rs\.\)|Payment \(Rs\.\))_\d+$", k):
                        all_keys.discard(k)

            final_headers = order_output_headers([k for k in all_keys if not is_dropped(k)])

            if sheet.max_column > len(final_headers):
                sheet.delete_cols(len(final_headers) + 1, sheet.max_column - len(final_headers))
            if sheet.max_row > len(row_records) + 1:
                sheet.delete_rows(len(row_records) + 2, sheet.max_row - (len(row_records) + 1))

            for col_idx, h in enumerate(final_headers, 1):
                sheet.cell(1, col_idx).value = h

            for row_idx, record in enumerate(row_records, 2):
                for col_idx, h in enumerate(final_headers, 1):
                    val = record.get(h)
                    if h == "Remarks":
                        raw_val = val if val is not None else record.get("scrape_status")
                        if raw_val is not None:
                            val_str = str(raw_val).strip().lower()
                            val = "Fetched" if val_str in ("fetched", "processed", "ok", "done") else "Bill Not Found"
                        else:
                            val = ""
                    elif val is not None and should_be_numeric(h):
                        val = excel_numeric_value(val)
                    sheet.cell(row_idx, col_idx).value = val if val is not None else ""

            if hasattr(sheet, "auto_filter"):
                max_c = len(final_headers)
                max_r = len(row_records) + 1
                sheet.auto_filter.ref = f"A1:{excel_col(max_c)}{max_r}"

            workbook.save(path)


class BillScraperApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PITC Bill Scrapper & MDI Analyzer")
        try:
            self.iconbitmap(resource_path("assets/mdiscrapper.ico"))
        except tk.TclError:
            pass
        self.geometry("640x630")
        self.minsize(580, 550)

        self.style = ttk.Style(self)
        try:
            self.style.theme_use("clam")
        except Exception:
            pass
        self.style.configure(
            "Green.Horizontal.TProgressbar",
            troughcolor="#e2e8f0",
            background="#16a34a",
            lightcolor="#16a34a",
            darkcolor="#15803d",
            bordercolor="#cbd5e1",
        )

        self.events = queue.Queue()
        self.stop_event = threading.Event()
        self.input_var = tk.StringVar()
        self.url_var = tk.StringVar(value=self._load_url())
        self.delay_var = tk.DoubleVar(value=2.0)
        self.history_1yr_var = tk.BooleanVar(value=False)
        self.history_6mo_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Ready. Choose a workbook (.xlsx/.ods) containing AC numbers in Column A.")
        self.total_var = tk.StringVar(value="0")
        self.ok_var = tk.StringVar(value="0")
        self.err_var = tk.StringVar(value="0")
        self.success_count = 0
        self.error_count = 0

        self.build_ui()
        self.after(100, self.drain_events)

    def _load_url(self):
        try:
            from pitcsuite.config import config_path as _cfg_p
            import configparser
            path = _cfg_p()
            if os.path.exists(path):
                cfg = configparser.ConfigParser()
                cfg.read(path)
                val = cfg.get("bill_scraper", "url", fallback="").strip()
                if val:
                    return val
        except Exception:
            pass
        return URL

    def _save_url(self, url):
        url = (url or "").strip()
        if not url:
            return
        try:
            from pitcsuite.config import config_path as _cfg_p
            import configparser
            path = _cfg_p()
            cfg = configparser.ConfigParser()
            if os.path.exists(path):
                cfg.read(path)
            if not cfg.has_section("bill_scraper"):
                cfg.add_section("bill_scraper")
            cfg.set("bill_scraper", "url", url)
            with open(path, "w", encoding="utf-8") as f:
                cfg.write(f)
        except Exception:
            pass

    def _toggle_url(self):
        if self.url_btn.cget("text") == "Edit":
            self.url_entry.configure(state="normal")
            self.url_btn.configure(text="Save")
            self.url_entry.focus()
        else:
            self.url_entry.configure(state="readonly")
            self.url_btn.configure(text="Edit")
            self._save_url(self.url_var.get().strip())

    def _on_toggle_history_1yr(self):
        if self.history_1yr_var.get():
            self.history_6mo_var.set(False)

    def _on_toggle_history_6mo(self):
        if self.history_6mo_var.get():
            self.history_1yr_var.set(False)

    def build_ui(self):
        root = ttk.Frame(self, padding=(18, 14, 18, 14))
        root.pack(fill="both", expand=True)

        header_frame = ttk.Frame(root)
        header_frame.pack(fill="x", pady=(0, 12))

        title_lbl = ttk.Label(header_frame, text="PITC Bill Scrapper & MDI Analyzer", font=("Segoe UI", 16, "bold"), foreground="#1e3a8a")
        title_lbl.pack(anchor="w")

        sub_lbl = ttk.Label(header_frame, text="Developed by Tahir Naqash | Automated Bill Scraping & Fixed Charges Audit", font=("Segoe UI", 9), foreground="#64748b")
        sub_lbl.pack(anchor="w", pady=(2, 0))

        files_frame = ttk.LabelFrame(root, text=" Target Workbook ", padding=10)
        files_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(files_frame, text="Workbook (.xlsx / .ods):", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", pady=4)
        entry = ttk.Entry(files_frame, textvariable=self.input_var, font=("Segoe UI", 9))
        entry.grid(row=0, column=1, sticky="ew", padx=(8, 8), pady=4)

        browse_btn = ttk.Button(files_frame, text="Browse...", command=self.choose_input)
        browse_btn.grid(row=0, column=2, pady=4)

        ttk.Label(files_frame, text="Bills URL:", font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky="w", pady=4)
        self.url_entry = ttk.Entry(files_frame, textvariable=self.url_var, font=("Segoe UI", 9), state="readonly")
        self.url_entry.grid(row=1, column=1, sticky="ew", padx=(8, 8), pady=4)
        self.url_entry.bind("<Return>", lambda e: self._toggle_url() if self.url_btn.cget("text") == "Save" else None)

        self.url_btn = ttk.Button(files_frame, text="Edit", command=self._toggle_url)
        self.url_btn.grid(row=1, column=2, pady=4)

        files_frame.columnconfigure(1, weight=1)

        mid_frame = ttk.Frame(root)
        mid_frame.pack(fill="x", pady=(0, 10))

        settings_box = ttk.LabelFrame(mid_frame, text=" Settings ", padding=8)
        settings_box.pack(side="left", fill="both", expand=True, padx=(0, 6))

        ttk.Label(settings_box, text="Delay (seconds):", font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", pady=2)
        spin = ttk.Spinbox(settings_box, from_=0.5, to=30.0, increment=0.5, textvariable=self.delay_var, width=6)
        spin.grid(row=0, column=1, sticky="w", padx=(6, 4), pady=2)
        ttk.Label(settings_box, text="(Default 2.0s prevents PITC blocking)", font=("Segoe UI", 8, "italic"), foreground="#64748b").grid(row=1, column=0, columnspan=2, sticky="w", pady=(2, 4))

        self.chk_history_1yr = ttk.Checkbutton(
            settings_box,
            text="Billing/Payment History (1 year)",
            variable=self.history_1yr_var,
            command=self._on_toggle_history_1yr,
        )
        self.chk_history_1yr.grid(row=2, column=0, columnspan=2, sticky="w", pady=(2, 1))

        self.chk_history_6mo = ttk.Checkbutton(
            settings_box,
            text="Billing Payment History (6 months)",
            variable=self.history_6mo_var,
            command=self._on_toggle_history_6mo,
        )
        self.chk_history_6mo.grid(row=3, column=0, columnspan=2, sticky="w", pady=(1, 2))

        stats_box = ttk.LabelFrame(mid_frame, text=" Progress Summary ", padding=8)
        stats_box.pack(side="right", fill="both", expand=True, padx=(6, 0))

        ttk.Label(stats_box, text="Total:").grid(row=0, column=0, sticky="w", padx=4)
        ttk.Label(stats_box, textvariable=self.total_var, font=("Segoe UI", 9, "bold")).grid(row=0, column=1, sticky="w", padx=4)

        ttk.Label(stats_box, text="Fetched:").grid(row=0, column=2, sticky="w", padx=4)
        ttk.Label(stats_box, textvariable=self.ok_var, font=("Segoe UI", 9, "bold"), foreground="#15803d").grid(row=0, column=3, sticky="w", padx=4)

        ttk.Label(stats_box, text="Errors:").grid(row=0, column=4, sticky="w", padx=4)
        ttk.Label(stats_box, textvariable=self.err_var, font=("Segoe UI", 9, "bold"), foreground="#b91c1c").grid(row=0, column=5, sticky="w", padx=4)

        self.progress = ttk.Progressbar(root, mode="determinate", style="Green.Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=(4, 6))

        status_lbl = ttk.Label(root, textvariable=self.status_var, font=("Segoe UI", 9), foreground="#0f172a")
        status_lbl.pack(anchor="w", pady=(0, 6))

        log_frame = ttk.LabelFrame(root, text=" Live Execution Log ", padding=6)
        log_frame.pack(fill="both", expand=True, pady=(0, 10))

        self.log = tk.Text(log_frame, height=10, state="disabled", wrap="word", font=("Consolas", 9), bg="#f8fafc", fg="#0f172a")
        self.log.tag_config("ok", foreground="#16a34a")
        self.log.tag_config("err", foreground="#dc2626")
        self.log.tag_config("info", foreground="#2563eb")
        self.log.tag_config("bold", font=("Consolas", 9, "bold"))

        scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=scroll.set)
        self.log.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        btn_bar = ttk.Frame(root)
        btn_bar.pack(fill="x")

        self.start_button = tk.Button(
            btn_bar, text="▶  Start Scraping", command=self.start,
            bg="#16a34a", fg="white", activebackground="#15803d", activeforeground="white",
            font=("Segoe UI", 10, "bold"), relief="raised", padx=16, pady=4, cursor="hand2"
        )
        self.start_button.pack(side="left")

        self.cancel_button = tk.Button(
            btn_bar, text="⏹  Stop", command=self.cancel,
            bg="#dc2626", fg="white", activebackground="#b91c1c", activeforeground="white",
            font=("Segoe UI", 10, "bold"), relief="raised", padx=14, pady=4, state="disabled", cursor="hand2"
        )
        self.cancel_button.pack(side="left", padx=(10, 0))

        self.open_btn = ttk.Button(btn_bar, text="Open File Folder", command=self.open_output_folder)
        self.open_btn.pack(side="right")

    def choose_input(self):
        path = filedialog.askopenfilename(
            title="Choose Account Workbook",
            filetypes=[("Supported Spreadsheets", "*.xlsx;*.ods"), ("Excel Workbook", "*.xlsx"), ("OpenDocument Spreadsheet", "*.ods"), ("All Files", "*.*")]
        )
        if path:
            self.input_var.set(path)
            self.status_var.set(f"Selected: {Path(path).name}")

    def open_output_folder(self):
        val = self.input_var.get().strip()
        if val and Path(val).parent.exists():
            folder = str(Path(val).parent.resolve())
            if sys.platform == "win32":
                os.startfile(folder)
            else:
                subprocess.Popen(["explorer", folder])
        else:
            messagebox.showinfo("Select File", "Please select a workbook file first.")

    def write_log(self, text, tag=None):
        self.log.configure(state="normal")
        if tag:
            self.log.insert("end", text + "\n", tag)
        else:
            self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def start(self):
        path = Path(self.input_var.get().strip())
        if not path.is_file():
            messagebox.showerror("Input Required", "Choose an existing workbook first.")
            return

        if not messagebox.askyesno(
            "Update Workbook",
            "The workbook will be updated in place with bill data, separate unit columns, cumulative MDI values, differences, Actual MDI, and fixed charges calculations.\n\nContinue?"
        ):
            return

        try:
            delay = max(0.1, float(self.delay_var.get()))
        except (TypeError, ValueError):
            messagebox.showerror("Invalid Delay", "Delay must be a positive number.")
            return

        history_mode = None
        if self.history_1yr_var.get():
            history_mode = "1yr"
        elif self.history_6mo_var.get():
            history_mode = "6mo"

        self.stop_event.clear()
        self.progress.configure(value=0)
        self.success_count = 0
        self.error_count = 0
        self.ok_var.set("0")
        self.err_var.set("0")
        self.total_var.set("0")

        self.start_button.configure(state="disabled")
        self.cancel_button.configure(state="normal")
        self.chk_history_1yr.configure(state="disabled")
        self.chk_history_6mo.configure(state="disabled")
        self.write_log(f"--- Starting Scraping on: {path.name} ---", "info")
        self.write_log(f"Request delay: {delay}s | Calculating Actual MDI (max diff x MF) & Actual Fix Charges...", "info")
        if history_mode == "1yr":
            self.write_log("Billing/Payment History: 1 Year (11 historical months) enabled", "info")
        elif history_mode == "6mo":
            self.write_log("Billing Payment History: 6 Months (5 historical months) enabled", "info")

        url = self.url_var.get().strip()
        self._save_url(url)
        if self.url_btn.cget("text") == "Save":
            self.url_entry.configure(state="readonly")
            self.url_btn.configure(text="Edit")

        threading.Thread(target=self.run, args=(path, delay, history_mode, url), daemon=True).start()

    def run(self, path, delay, history_mode=None, url=None):
        try:
            def progress(done, total, row):
                raw_status = row.get("Remarks") or row.get("scrape_status", "")
                if str(raw_status).strip().lower() in ("ok", "done", "processed", "fetched"):
                    status_text = "Fetched"
                else:
                    status_text = "Bill Not Found"
                self.events.put(("progress", done, total, row.get("_account", ""), status_text))

            def log_cb(msg):
                self.events.put(("log", msg))

            PITCBillScraper(delay=delay, history_mode=history_mode, url=url).update_workbook(path, progress, self.stop_event, log_callback=log_cb)
            self.events.put(("done", str(path), self.stop_event.is_set()))
        except Exception as exc:
            self.events.put(("error", str(exc)))

    def cancel(self):
        self.stop_event.set()
        self.status_var.set("Finishing current bill, then stopping...")
        self.cancel_button.configure(state="disabled")
        self.write_log("Stop requested. Waiting for current account to finish...", "err")

    def drain_events(self):
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == "progress":
                    _, done, total, account, status = event
                    self.progress.configure(maximum=max(total, 1), value=done)
                    self.total_var.set(str(total))
                    display_status = "Fetched" if str(status).strip().lower() in ("ok", "done", "processed", "fetched") else "Bill Not Found"
                    self.status_var.set(f"Fetched {done} of {total} accounts...")

                    if display_status == "Fetched":
                        self.success_count += 1
                        self.ok_var.set(str(self.success_count))
                        self.write_log(f"[{done}/{total}] {account}: Fetched", "ok")
                    else:
                        self.error_count += 1
                        self.err_var.set(str(self.error_count))
                        self.write_log(f"[{done}/{total}] {account}: Bill Not Found", "err")

                elif event[0] == "log":
                    self.write_log(event[1], "info")

                elif event[0] == "done":
                    _, path, cancelled = event
                    self.start_button.configure(state="normal")
                    self.cancel_button.configure(state="disabled")
                    self.chk_history_1yr.configure(state="normal")
                    self.chk_history_6mo.configure(state="normal")
                    self.status_var.set("Stopped by user (All progress saved)." if cancelled else "Completed successfully!")
                    self.write_log(
                        f"=== {'Scraping Stopped. All progress saved to:' if cancelled else 'All Done! Successfully updated:'} {path} ===",
                        "info" if not cancelled else "err"
                    )
                    if not cancelled:
                        messagebox.showinfo(
                            "Processing Finished",
                            f"Workbook updated successfully with separated unit columns, Actual MDI, and Fixed Charges calculations:\n\n{path}"
                        )

                elif event[0] == "error":
                    self.start_button.configure(state="normal")
                    self.cancel_button.configure(state="disabled")
                    self.chk_history_1yr.configure(state="normal")
                    self.chk_history_6mo.configure(state="normal")
                    self.status_var.set("Failed with error.")
                    self.write_log("ERROR: " + event[1], "err")
                    messagebox.showerror("Export Failed", event[1])
        except queue.Empty:
            pass
        self.after(100, self.drain_events)


if __name__ == "__main__":
    BillScraperApp().mainloop()

"""
CP136 Change Posting Report Engine for PITC Suite
Parses WAPDA RCP-136 change posting reports (.txt), extracts 14-digit Ref Nos,
handles deduplication (single postings), and exports formatted Excel (.xlsx) files.
"""

import os
import re
import glob
from typing import List, Dict, Any, Tuple, Optional, Callable
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# Ordered list of all standard headings in WAPDA RCP-136 reports
HEADINGS_TOP = [
    'APPLICATION NO', 'APP DATE', 'TARIFF', 'SUB BATCH', 'SEC AMOUNT', 'SEC DATE'
]

HEADINGS_BLOCK_1 = [
    'CONN DATE', 'SAN LOAD', 'RU', 'F/P', 'SEASON CD', 'ED CD',
    'GVT DEP', 'STD CLSF', 'IT', 'METERS', 'AC', 'NTN', 'GST'
]

HEADINGS_BLOCK_2 = [
    'METER RENT', 'SRV RENT', 'N .ID NUMBER', 'TRANSFRMR CODE',
    'CONNECTION CATEGORY CODE', 'NET MTR CD', 'NET MTR CAP',
    'NET MTR CONNECTION DATE', 'NET MTR EXPIRY DATE'
]

HEADINGS_BLOCK_3 = [
    'SET NO', 'METER TYPE', 'FEEDER CODE', 'METER NO', 'LOC CODE',
    'MFG CODE', 'METER CODE', 'M/PLY FACTOR', 'MTR RANGE',
    'METER READING', 'ERRORS CODES'
]

ALL_ORDERED_HEADINGS = HEADINGS_TOP + HEADINGS_BLOCK_1 + HEADINGS_BLOCK_2 + HEADINGS_BLOCK_3

# Character slice boundaries for fixed-width columns in each block
SPANS_BLOCK_1: List[Tuple[str, int, int]] = [
    ('CONN DATE', 0, 14),
    ('SAN LOAD', 14, 27),
    ('RU', 27, 34),
    ('F/P', 34, 42),
    ('SEASON CD', 42, 55),
    ('ED CD', 55, 64),
    ('GVT DEP', 64, 73),
    ('STD CLSF', 73, 82),
    ('IT', 82, 87),
    ('METERS', 87, 97),
    ('AC', 97, 103),
    ('NTN', 103, 115),
    ('GST', 115, 135),
]

SPANS_BLOCK_2: List[Tuple[str, int, int]] = [
    ('METER RENT', 0, 15),
    ('SRV RENT', 15, 28),
    ('N .ID NUMBER', 28, 44),
    ('TRANSFRMR CODE', 44, 60),
]

SPANS_BLOCK_3: List[Tuple[str, int, int]] = [
    ('SET NO', 0, 9),
    ('METER TYPE', 9, 21),
    ('FEEDER CODE', 21, 31),
    ('METER NO', 31, 40),
    ('LOC CODE', 40, 49),
    ('MFG CODE', 49, 58),
    ('METER CODE', 58, 66),
    ('M/PLY FACTOR', 66, 76),
    ('MTR RANGE', 76, 83),
    ('METER READING', 83, 97),
    ('ERRORS CODES', 97, 135),
]


def clean_page_banners(raw_text: str) -> List[str]:
    """
    Splits text across form-feed characters and strips standard page header banners.
    """
    pages = raw_text.split('\x0c') if '\x0c' in raw_text else [raw_text]
    unified_lines: List[str] = []

    for page in pages:
        lines = page.splitlines()
        banner_end = -1
        for idx, l in enumerate(lines[:15]):
            stripped = l.strip()
            if stripped and set(stripped) == {'-'}:
                banner_end = idx
                break

        if banner_end != -1:
            page_content = lines[banner_end + 1:]
        else:
            page_content = lines

        for line in page_content:
            if any(marker in line for marker in [
                'WAPDA ELECTRICITY BILLING',
                'MASTER DATA CHANGE',
                'POSTING LIST FOR MAIN DATE',
                'DIVISON TOTAL:',
                'GRAND TOTAL:',
                'Total Records For the MainDate',
                'Total Records Posted',
                'Total Records UnPosted',
            ]):
                continue
            unified_lines.append(line)

    return unified_lines


def parse_single_txt_file(filepath: str) -> List[Dict[str, Any]]:
    """
    Parses a single CP-136 change posting text file and extracts structured records.
    Each record contains 'Ref No', 'Status', and values for modified columns.
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"File not found: {filepath}")

    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
        raw_text = f.read()

    unified_lines = clean_page_banners(raw_text)

    ref_pattern = re.compile(r'^\s*(\d{14})')
    err_pattern = re.compile(r'\b([A-Z0-9]{2,5},)')

    ref_indices: List[int] = []
    for idx, line in enumerate(unified_lines):
        if ref_pattern.match(line):
            ref_indices.append(idx)

    records: List[Dict[str, Any]] = []
    base_filename = os.path.basename(filepath)

    ignored_b1 = ['CONN DATE', 'TOTAL', '----------------']
    ignored_b2 = [
        'METER RENT', 'CONTACT NO', 'CONSUMER NAME', 'CONNECTION CATEGORY',
        'FATHER/HUSBAND', 'ADDRES LINE', 'TRANSFORMER', 'NET MTR', 'SIM ID', 'SIM NO',
        'SET NO', '----------------'
    ]
    ignored_b3 = [
        'SET NO', 'CODE      NO', '----------------', 'New Ref No',
        'POSTED', 'NOT POSTED'
    ]

    for idx, r_i in enumerate(ref_indices):
        next_i = ref_indices[idx + 1] if idx + 1 < len(ref_indices) else len(unified_lines)
        rec_lines = unified_lines[r_i:next_i]

        first_line = rec_lines[0]
        m_ref = ref_pattern.match(first_line)
        ref_no = m_ref.group(1) if m_ref else ''

        after_ref = first_line.lstrip()[14:].strip()
        customer_id = ''
        if after_ref:
            cust_match = re.match(r'^(\d+)', after_ref)
            if cust_match:
                customer_id = cust_match.group(1)

        # Determine POSTED / NOT POSTED status
        status = ''
        for line in rec_lines:
            if 'NOT POSTED' in line:
                status = 'NOT POSTED'
                break
            elif 'POSTED' in line and 'Total Records' not in line:
                status = 'POSTED'

        # Extract error codes
        error_codes: List[str] = []
        for line in rec_lines:
            if 'Ref No' in line or 'SET NO' in line or 'ERROR CODES' in line or 'ERRORS CODES' in line:
                continue
            if len(line) > 95:
                found = err_pattern.findall(line[95:])
                for code in found:
                    code_clean = code.strip(',')
                    if code_clean not in error_codes and code_clean not in ['POSTED', 'TOTAL']:
                        error_codes.append(code_clean)

        record_data: Dict[str, Any] = {
            'Ref No': ref_no,
            'Status': status,
            'Customer ID': customer_id,
            'File Name': base_filename,
        }

        if error_codes:
            record_data['ERRORS CODES'] = ', '.join(error_codes)

        # 1. Parse Block 1 (CONN DATE ... GST)
        for j, line in enumerate(rec_lines):
            if 'CONN DATE' in line:
                for k in range(j + 1, len(rec_lines)):
                    if 'METER RENT' in rec_lines[k] or '----------------' in rec_lines[k]:
                        break
                    s = rec_lines[k].strip()
                    if not s or any(ign in s for ign in ignored_b1):
                        continue
                    for col_name, start, end in SPANS_BLOCK_1:
                        val = rec_lines[k][start:end].strip() if start < len(rec_lines[k]) else ''
                        if val:
                            record_data[col_name] = val
                break

        # 2. Parse Block 2 (METER RENT ... TRANSFRMR CODE)
        for j, line in enumerate(rec_lines):
            if 'METER RENT' in line:
                for k in range(j + 1, len(rec_lines)):
                    if 'SET NO' in rec_lines[k] or 'CONSUMER NAME' in rec_lines[k]:
                        break
                    s = rec_lines[k].strip()
                    if not s or any(ign in rec_lines[k] for ign in ignored_b2):
                        continue
                    for col_name, start, end in SPANS_BLOCK_2:
                        val = rec_lines[k][start:end].strip() if start < len(rec_lines[k]) else ''
                        if val:
                            record_data[col_name] = val
                break

        # 2b. Net Metering checks in rec_lines
        for line in rec_lines:
            if 'NET MTR CD' in line or 'NET MTR C' in line:
                m_cap = re.search(r'NET MTR\s+(?:CAP|C)\s+([0-9.]+)', line)
                if m_cap:
                    record_data['NET MTR CAP'] = m_cap.group(1)
                m_cd = re.search(r'NET MTR\s+CD\s+([0-9A-Z]+)', line)
                if m_cd:
                    record_data['NET MTR CD'] = m_cd.group(1)
            if 'NET MTR CONNECTION DAT' in line or 'NET MTR CONNECTION DATE' in line:
                m_conn = re.search(r'NET MTR CONNECTION DAT[E]?\s+([0-9/:\sA-Za-z]+)', line)
                if m_conn:
                    record_data['NET MTR CONNECTION DATE'] = m_conn.group(1).strip()
            if 'NET MTR EXPIRY DA' in line or 'NET MTR EXPIRY DATE' in line:
                m_exp = re.search(r'NET MTR EXPIRY DA[TE]?\s+([0-9/:\sA-Za-z]+)', line)
                if m_exp:
                    record_data['NET MTR EXPIRY DATE'] = m_exp.group(1).strip()

        # 3. Parse Block 3 (SET NO ... METER READING)
        set_lines: List[str] = []
        found_set = False
        for line in rec_lines:
            if 'SET NO' in line:
                found_set = True
                continue
            if found_set:
                if 'New Ref No' in line or 'POSTED' in line or 'Ref No' in line:
                    break
                s = line.strip()
                if not s or any(ign in s for ign in ignored_b3):
                    continue
                set_lines.append(line)

        if set_lines:
            for sline in set_lines:
                row_copy = dict(record_data)
                for col_name, start, end in SPANS_BLOCK_3:
                    if col_name == 'ERRORS CODES':
                        continue
                    val = sline[start:end].strip() if start < len(sline) else ''
                    if val:
                        row_copy[col_name] = val
                records.append(row_copy)
        else:
            records.append(record_data)

    return records


def parse_single_pdf_file(filepath: str) -> List[Dict[str, Any]]:
    """
    Parses a single CP-136 change posting PDF report and extracts structured records.
    Accurately extracts 14-figure Ref Nos, Status, Net Metering changes, Meter Sets,
    and field modifications.
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"File not found: {filepath}")

    try:
        import pymupdf as fitz
    except ImportError:
        import fitz

    doc = fitz.open(filepath)
    records: List[Dict[str, Any]] = []
    base_filename = os.path.basename(filepath)
    err_pattern = re.compile(r'\b([A-Z0-9]{2,5},)')

    for page_idx, page in enumerate(doc):
        words = page.get_text("words")
        if not words:
            continue

        # Look for 14-digit Ref No words (in left margin x < 120)
        ref_words = [w for w in words if re.match(r'^\d{14}$', w[4]) and w[0] < 120]
        if not ref_words:
            continue

        page_text = page.get_text("text")

        # Determine landmark y-coordinates
        lines_dict: Dict[float, List[Any]] = {}
        for w in words:
            y = round(w[1] / 3.5) * 3.5
            lines_dict.setdefault(y, []).append(w)

        sorted_y = sorted(lines_dict.keys())
        y_b1_hdr = None
        y_b2_hdr = None
        y_b3_hdr = None
        y_status = None

        for y in sorted_y:
            lw = lines_dict[y]
            line_str = " ".join([w[4] for w in sorted(lw, key=lambda x: x[0])])
            if 'CONN DATE' in line_str or ('CONN' in line_str and 'SAN' in line_str):
                y_b1_hdr = y
            elif 'METER RENT' in line_str or ('SRV RENT' in line_str and 'TRANSFRMR' in line_str):
                y_b2_hdr = y
            elif 'SET NO' in line_str or ('SET' in line_str and 'FEEDER' in line_str):
                y_b3_hdr = y
            elif 'POSTED' in line_str or 'NOT POSTED' in line_str:
                y_status = y

        # Status
        status = ''
        if 'NOT POSTED' in page_text:
            status = 'NOT POSTED'
        elif 'POSTED' in page_text:
            status = 'POSTED'

        # Extract error codes
        errs = []
        for e in err_pattern.findall(page_text):
            code = e.strip(',')
            if code not in errs and code not in ['POSTED', 'TOTAL']:
                errs.append(code)

        # Process each Ref No found on this page
        sorted_refs = sorted(ref_words, key=lambda x: x[1])

        for r_idx, rw in enumerate(sorted_refs):
            ref_no = rw[4]
            y_curr_ref = rw[1]
            y_next_ref = sorted_refs[r_idx + 1][1] if r_idx + 1 < len(sorted_refs) else 9999.0

            rec_words = [w for w in words if y_curr_ref - 10 <= w[1] < y_next_ref - 10]

            # Customer ID
            cust_id = ''
            c_words = [w for w in rec_words if abs(w[1] - y_curr_ref) < 8 and 100 < w[0] < 200 and w[4].isdigit()]
            if c_words:
                cust_id = c_words[0][4]

            record_data: Dict[str, Any] = {
                'Ref No': ref_no,
                'Status': status,
                'Customer ID': cust_id,
                'File Name': base_filename
            }
            if errs:
                record_data['ERRORS CODES'] = ", ".join(errs)

            # Top block values (TARIFF, APPLICATION NO, APP DATE, SEC AMOUNT, SEC DATE, SUB BATCH)
            top_cols = [
                ('SUB BATCH', 200, 270),
                ('TARIFF', 270, 315),
                ('APPLICATION NO', 315, 410),
                ('APP DATE', 410, 470),
                ('SEC AMOUNT', 470, 550),
                ('SEC DATE', 550, 620),
            ]
            top_words = [w for w in rec_words if abs(w[1] - y_curr_ref) < 8 and w[0] >= 200]
            for col_name, x0, x1 in top_cols:
                cw = [w[4] for w in top_words if x0 <= w[0] < x1]
                if cw:
                    record_data[col_name] = " ".join(cw)

            # Block 1 values: between y_b1_hdr and separator / y_b2_hdr
            if y_b1_hdr:
                b1_cols = [
                    ('CONN DATE', 25, 115),
                    ('SAN LOAD', 115, 180),
                    ('RU', 180, 235),
                    ('F/P', 235, 275),
                    ('SEASON CD', 275, 350),
                    ('ED CD', 350, 400),
                    ('GVT DEP', 400, 455),
                    ('STD CLSF', 455, 510),
                    ('IT', 510, 545),
                    ('METERS', 545, 605),
                    ('AC', 605, 640),
                    ('NTN', 640, 720),
                    ('GST', 720, 820),
                ]
                b1_words = [w for w in rec_words if y_b1_hdr + 5 < w[1] < (y_b2_hdr if y_b2_hdr else y_b1_hdr + 35) and not w[4].startswith('---')]
                for col_name, x0, x1 in b1_cols:
                    cw = [w[4] for w in b1_words if x0 <= w[0] < x1]
                    if cw:
                        record_data[col_name] = " ".join(cw)

            # Block 2 values (METER RENT ... TRANSFRMR CODE)
            if y_b2_hdr:
                b2_cols = [
                    ('METER RENT', 25, 115),
                    ('SRV RENT', 115, 185),
                    ('N .ID NUMBER', 185, 275),
                    ('TRANSFRMR CODE', 275, 380),
                ]
                b2_top_words = [w for w in rec_words if y_b2_hdr + 5 < w[1] < y_b2_hdr + 30 and not w[4].startswith('---')]
                for col_name, x0, x1 in b2_cols:
                    cw = [w[4] for w in b2_top_words if x0 <= w[0] < x1]
                    if cw:
                        record_data[col_name] = " ".join(cw)

                # Net Metering & Connection Category in Block 2
                net_cd = [w[4] for w in rec_words if 288 <= w[1] <= 304 and 420 <= w[0] <= 460]
                if net_cd:
                    record_data['NET MTR CD'] = " ".join(net_cd)

                net_cap = [w[4] for w in rec_words if 288 <= w[1] <= 304 and 580 <= w[0] <= 630]
                if net_cap:
                    record_data['NET MTR CAP'] = " ".join(net_cap)

                net_conn = [w[4] for w in rec_words if 236 <= w[1] <= 246 and 680 <= w[0] <= 800]
                if net_conn:
                    record_data['NET MTR CONNECTION DATE'] = " ".join(net_conn)

                net_exp = [w[4] for w in rec_words if 255 <= w[1] <= 265 and 680 <= w[0] <= 800]
                if net_exp:
                    record_data['NET MTR EXPIRY DATE'] = " ".join(net_exp)

            # Block 3 values (SET NO ... METER READING)
            b3_cols = [
                ('SET NO', 25, 75),
                ('METER TYPE', 75, 140),
                ('FEEDER CODE', 140, 200),
                ('METER NO', 200, 260),
                ('LOC CODE', 260, 315),
                ('MFG CODE', 315, 365),
                ('METER CODE', 365, 410),
                ('M/PLY FACTOR', 410, 470),
                ('MTR RANGE', 470, 505),
                ('METER READING', 505, 575),
            ]

            b3_lines_dict = {}
            if y_b3_hdr:
                for w in rec_words:
                    if y_b3_hdr + 15 < w[1] < (y_status if y_status else y_b3_hdr + 65):
                        if not w[4].startswith('---') and w[4] not in ['New', 'Ref', 'No', 'POSTED', 'NOT']:
                            ly = round(w[1] / 4.0) * 4.0
                            b3_lines_dict.setdefault(ly, []).append(w)

            if b3_lines_dict:
                for ly in sorted(b3_lines_dict.keys()):
                    lw = b3_lines_dict[ly]
                    r = dict(record_data)
                    for col_name, x0, x1 in b3_cols:
                        cw = [w[4] for w in lw if x0 <= w[0] < x1]
                        if cw:
                            r[col_name] = " ".join(cw)
                    records.append(r)
            else:
                records.append(record_data)

    return records


def parse_single_file(filepath: str) -> List[Dict[str, Any]]:
    """
    Parses a single CP-136 change posting report (.txt or .pdf) and extracts structured records.
    Each record contains 'Ref No', 'Status', and values for modified columns.
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"File not found: {filepath}")

    if filepath.lower().endswith('.pdf'):
        return parse_single_pdf_file(filepath)
    return parse_single_txt_file(filepath)


def deduplicate_records(
    records: List[Dict[str, Any]],
    mode: str = 'prefer_posted'
) -> Tuple[List[Dict[str, Any]], int]:
    """
    Deduplicates records by 14-figure Reference Number ('Ref No').
    Ensures that only single postings remain in the output, deleting 2nd and duplicate findings.

    Args:
        records: List of parsed record dicts.
        mode:
            'prefer_posted': Keeps 'POSTED' if any occurrence succeeded, otherwise keeps 1st occurrence.
                             Merges any non-empty field changes across occurrences.
            'first': Strictly keeps the 1st finding and deletes 2nd and subsequent findings.
            'last': Keeps the last finding and deletes earlier findings.

    Returns:
        (unique_records, count_of_duplicates_removed)
    """
    if not records:
        return [], 0

    seen: Dict[str, Dict[str, Any]] = {}
    order: List[str] = []
    duplicates_removed = 0

    for rec in records:
        ref = rec.get('Ref No', '')
        if not ref:
            continue

        if ref not in seen:
            seen[ref] = dict(rec)
            order.append(ref)
        else:
            duplicates_removed += 1
            existing = seen[ref]

            if mode == 'first':
                continue

            elif mode == 'last':
                seen[ref] = dict(rec)

            else:  # 'prefer_posted' (default)
                if rec.get('Status') == 'POSTED' and existing.get('Status') != 'POSTED':
                    existing['Status'] = 'POSTED'
                    existing['Customer ID'] = rec.get('Customer ID', existing.get('Customer ID'))
                    existing['ERRORS CODES'] = rec.get('ERRORS CODES', '')
                    existing['File Name'] = rec.get('File Name', existing.get('File Name'))

                for k, v in rec.items():
                    if v and not existing.get(k):
                        existing[k] = v

    unique_records = [seen[ref] for ref in order]
    return unique_records, duplicates_removed


def parse_multiple_files(
    filepaths: List[str],
    deduplicate: bool = True,
    dedup_mode: str = 'prefer_posted'
) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """
    Parses multiple CP-136 files and aggregates their records.
    Optionally deduplicates by Reference Number to keep single postings.

    Returns:
        all_records: list of parsed (and optionally deduplicated) records
        summary_stats: summary counts
    """
    raw_records: List[Dict[str, Any]] = []
    stats = {
        'total_files': len(filepaths),
        'raw_records': 0,
        'duplicates_removed': 0,
        'total_records': 0,
        'posted_records': 0,
        'not_posted_records': 0,
        'files_summary': {}
    }

    for path in filepaths:
        try:
            records = parse_single_file(path)
            posted = sum(1 for r in records if r.get('Status') == 'POSTED')
            not_posted = sum(1 for r in records if r.get('Status') == 'NOT POSTED')
            raw_records.extend(records)

            stats['files_summary'][os.path.basename(path)] = {
                'records': len(records),
                'posted': posted,
                'not_posted': not_posted,
                'status': 'OK'
            }
        except Exception as e:
            stats['files_summary'][os.path.basename(path)] = {
                'records': 0,
                'posted': 0,
                'not_posted': 0,
                'status': f'Error: {str(e)}'
            }

    stats['raw_records'] = len(raw_records)

    if deduplicate:
        final_records, dups_removed = deduplicate_records(raw_records, mode=dedup_mode)
        stats['duplicates_removed'] = dups_removed
    else:
        final_records = raw_records
        stats['duplicates_removed'] = 0

    stats['total_records'] = len(final_records)
    stats['posted_records'] = sum(1 for r in final_records if r.get('Status') == 'POSTED')
    stats['not_posted_records'] = sum(1 for r in final_records if r.get('Status') == 'NOT POSTED')

    return final_records, stats


def export_to_excel(
    records: List[Dict[str, Any]],
    output_path: str,
    only_present_columns: bool = True,
    include_status: bool = True,
    include_filename: bool = False,
    include_summary_sheet: bool = True,
    summary_stats: Dict[str, Any] = None
) -> str:
    """
    Exports parsed CP136 records to an Excel workbook.
    Formats Reference Numbers as text ('@') to prevent scientific notation.
    """
    wb = openpyxl.Workbook()

    # Determine which columns to include
    active_headings: List[str] = []
    if only_present_columns:
        for heading in ALL_ORDERED_HEADINGS:
            if any(rec.get(heading) for rec in records):
                active_headings.append(heading)
    else:
        active_headings = list(ALL_ORDERED_HEADINGS)

    columns = ['Ref No']
    if include_status:
        columns.append('Status')
    if include_filename:
        columns.append('File Name')
    columns.extend(active_headings)

    # Main data sheet
    ws_data = wb.active
    ws_data.title = 'CP136 Posting Data'
    ws_data.views.sheetView[0].showGridLines = True

    # Styling
    font_header = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
    fill_header = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')

    font_data = Font(name='Calibri', size=10)
    fill_even = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    fill_odd = PatternFill(start_color='F4F7FA', end_color='F4F7FA', fill_type='solid')

    fill_posted = PatternFill(start_color='D4EDDA', end_color='D4EDDA', fill_type='solid')
    font_posted = Font(name='Calibri', size=10, bold=True, color='155724')

    fill_not_posted = PatternFill(start_color='F8D7DA', end_color='F8D7DA', fill_type='solid')
    font_not_posted = Font(name='Calibri', size=10, bold=True, color='721C24')

    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')
    align_header = Alignment(horizontal='center', vertical='center', wrap_text=True)

    border_thin = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    border_header = Border(
        left=Side(style='thin', color='13324E'),
        right=Side(style='thin', color='13324E'),
        top=Side(style='thin', color='13324E'),
        bottom=Side(style='medium', color='0F2438')
    )

    # Write Header row
    ws_data.row_dimensions[1].height = 28
    for col_idx, col_name in enumerate(columns, start=1):
        cell = ws_data.cell(row=1, column=col_idx, value=col_name)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_header
        cell.border = border_header

    # Write Data rows
    for row_idx, rec in enumerate(records, start=2):
        ws_data.row_dimensions[row_idx].height = 20
        row_fill = fill_odd if (row_idx % 2 == 1) else fill_even

        for col_idx, col_name in enumerate(columns, start=1):
            val = rec.get(col_name, '')
            cell = ws_data.cell(row=row_idx, column=col_idx)

            if col_name == 'Ref No':
                cell.number_format = '@'
                cell.value = str(val) if val else ''
                cell.alignment = align_center
                cell.font = font_data
                cell.fill = row_fill
            elif col_name == 'Status':
                cell.value = str(val) if val else ''
                cell.alignment = align_center
                if val == 'POSTED':
                    cell.fill = fill_posted
                    cell.font = font_posted
                elif val == 'NOT POSTED':
                    cell.fill = fill_not_posted
                    cell.font = font_not_posted
                else:
                    cell.fill = row_fill
                    cell.font = font_data
            else:
                cell.value = val
                cell.alignment = align_center if len(str(val)) <= 15 else align_left
                cell.font = font_data
                cell.fill = row_fill

            cell.border = border_thin

    ws_data.freeze_panes = 'A2'
    ws_data.auto_filter.ref = ws_data.dimensions

    # Auto-adjust column widths
    for col_idx, col_name in enumerate(columns, start=1):
        col_letter = get_column_letter(col_idx)
        max_len = len(str(col_name))
        for row in range(2, min(len(records) + 2, 200)):
            cell_val = str(ws_data.cell(row=row, column=col_idx).value or '')
            if len(cell_val) > max_len:
                max_len = len(cell_val)
        ws_data.column_dimensions[col_letter].width = max(max_len + 4, 13)

    # Optional Summary Sheet
    if include_summary_sheet and summary_stats:
        ws_sum = wb.create_sheet(title='Summary', index=0)
        ws_sum.views.sheetView[0].showGridLines = True

        title_font = Font(name='Calibri', size=14, bold=True, color='1F4E78')
        sec_font = Font(name='Calibri', size=11, bold=True)
        lbl_font = Font(name='Calibri', size=11, bold=True, color='333333')
        val_font = Font(name='Calibri', size=11)

        ws_sum['A1'] = "CP-136 Change Posting Report Summary"
        ws_sum['A1'].font = title_font
        ws_sum.row_dimensions[1].height = 28

        stats_items = [
            ("Total Files Processed", summary_stats.get('total_files', 0)),
        ]
        if summary_stats.get('duplicates_removed', 0) > 0:
            stats_items.extend([
                ("Total Records Found (Raw)", summary_stats.get('raw_records', summary_stats.get('total_records', 0))),
                ("Duplicate Records Removed", summary_stats.get('duplicates_removed', 0)),
                ("Unique Single Records Exported", summary_stats.get('total_records', len(records))),
            ])
        else:
            stats_items.append(("Total Records Extracted", summary_stats.get('total_records', len(records))))

        stats_items.extend([
            ("Total Records Posted", summary_stats.get('posted_records', 0)),
            ("Total Records Unposted (Not Posted)", summary_stats.get('not_posted_records', 0)),
        ])

        for s_idx, (label, value) in enumerate(stats_items, start=3):
            c_lbl = ws_sum.cell(row=s_idx, column=1, value=label)
            c_lbl.font = lbl_font
            c_val = ws_sum.cell(row=s_idx, column=2, value=value)
            c_val.font = val_font
            c_val.alignment = align_center

        f_start = len(stats_items) + 5
        ws_sum.cell(row=f_start, column=1, value="Files Breakdown").font = sec_font

        f_hdr_row = f_start + 1
        ws_sum.row_dimensions[f_hdr_row].height = 24
        file_hdrs = ["File Name", "Records", "Posted", "Not Posted", "Status"]
        for f_c, h in enumerate(file_hdrs, start=1):
            cell = ws_sum.cell(row=f_hdr_row, column=f_c, value=h)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = align_center

        f_data_row = f_hdr_row + 1
        for fname, f_info in summary_stats.get('files_summary', {}).items():
            ws_sum.cell(row=f_data_row, column=1, value=fname).alignment = align_left
            ws_sum.cell(row=f_data_row, column=2, value=f_info.get('records', 0)).alignment = align_center
            ws_sum.cell(row=f_data_row, column=3, value=f_info.get('posted', 0)).alignment = align_center
            ws_sum.cell(row=f_data_row, column=4, value=f_info.get('not_posted', 0)).alignment = align_center
            ws_sum.cell(row=f_data_row, column=5, value=f_info.get('status', 'OK')).alignment = align_center
            f_data_row += 1

        ws_sum.column_dimensions['A'].width = 35
        ws_sum.column_dimensions['B'].width = 15
        ws_sum.column_dimensions['C'].width = 15
        ws_sum.column_dimensions['D'].width = 15
        ws_sum.column_dimensions['E'].width = 15

    # Ensure output directory exists
    out_dir = os.path.dirname(os.path.abspath(output_path))
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    wb.save(output_path)
    return os.path.abspath(output_path)


def run_cp136_check(
    file_paths: List[str],
    output_path: Optional[str] = None,
    deduplicate: bool = True,
    dedup_mode: str = 'prefer_posted',
    only_present: bool = True,
    include_status: bool = True,
    include_filename: bool = False,
    log_fn: Optional[Callable[[str], None]] = None
) -> Tuple[str, Dict[str, Any]]:
    """
    High-level runner used by PITC Suite Worker.
    Returns (saved_excel_path, stats_dict).
    """
    def _log(msg: str):
        if log_fn:
            log_fn(msg)

    if not file_paths:
        raise ValueError("Please select at least one CP-136 report file (.txt or .pdf).")

    _log(f"Processing {len(file_paths)} CP-136 file(s)...")

    records, stats = parse_multiple_files(
        filepaths=file_paths,
        deduplicate=deduplicate,
        dedup_mode=dedup_mode
    )

    if not records:
        raise ValueError("No records found in the selected CP-136 file(s).")

    _log(f"Found {stats.get('raw_records', len(records))} raw record(s).")
    if stats.get('duplicates_removed', 0) > 0:
        _log(f"Removed {stats['duplicates_removed']} duplicate / 2nd finding(s) (Single Postings Mode).")
    _log(f"Single postings ready for export: {len(records)} ({stats.get('posted_records', 0)} POSTED, {stats.get('not_posted_records', 0)} NOT POSTED)")

    # Resolve output destination
    if not output_path or not output_path.strip():
        first_dir = os.path.dirname(file_paths[0])
        first_base = os.path.splitext(os.path.basename(file_paths[0]))[0]
        suffix = "_Single_Postings.xlsx" if deduplicate else "_Extracted.xlsx"
        output_path = os.path.join(first_dir, f"{first_base}{suffix}")
    elif os.path.isdir(output_path.strip()):
        first_base = os.path.splitext(os.path.basename(file_paths[0]))[0]
        suffix = "_Single_Postings.xlsx" if deduplicate else "_Extracted.xlsx"
        output_path = os.path.join(output_path.strip(), f"{first_base}{suffix}")
    elif not output_path.strip().lower().endswith('.xlsx'):
        output_path = output_path.strip() + '.xlsx'
    else:
        output_path = output_path.strip()

    saved_path = export_to_excel(
        records=records,
        output_path=output_path,
        only_present_columns=only_present,
        include_status=include_status,
        include_filename=include_filename or len(file_paths) > 1,
        summary_stats=stats
    )

    _log(f"Successfully saved Excel file: {os.path.basename(saved_path)}")
    return saved_path, stats

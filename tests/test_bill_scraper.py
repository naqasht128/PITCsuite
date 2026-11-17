import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock
from openpyxl import Workbook
from bs4 import BeautifulSoup
from pitcsuite.tools.bill_scraper_logic import (
    PITCBillScraper,
    compute_total_units,
    compute_total_payment,
    is_header_text,
    is_ac_header,
    is_sr_header,
)

SAMPLE_HTML = """
<!DOCTYPE html>
<html>
<body>
<div class="gbn-app-shell">
    <div class="grid-col-cell">
        <div class="label-row">
            <span class="en-lbl">STATUS</span>
            <span class="ur-lbl">&#x062D;&#x0627;&#x0644;&#x062A;</span>
        </div>
        <div class="val-space">ALIVE</div>
    </div>
    <div class="grid-col-cell">
        <div class="label-row">
            <span class="en-lbl">NAME & ADDRESS</span>
            <span class="ur-lbl">&#x0646;&#x0627;&#x0645; &#x0648; &#x067e;&#x062a;&#x06c1;</span>
        </div>
        <div class="val-space">Edoto Mobile Tower, Near National Club, D.I.Khan</div>
    </div>
    <div class="grid-col-cell">
        <div class="label-row">
            <span class="en-lbl">DUE DATE</span>
            <span class="ur-lbl">&#x0645;&#x0642;&#x0631;&#x0631;&#x06c1; &#x062a;&#x0627;&#x0631;&#x06cc;&#x062e;</span>
        </div>
        <div class="val-space">23 OCT 26</div>
    </div>

    <table>
        <thead>
            <tr>
                <th>METER NO</th>
                <th>PREVIOUS READING</th>
                <th>PRESENT READING</th>
                <th>MF</th>
                <th>UNITS</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>3798029575</td>
                <td>25246.07<br>6494.01</td>
                <td>26205.61<br>6754.78</td>
                <td>1<br>1</td>
                <td>960<br>261</td>
            </tr>
        </tbody>
    </table>
</div>
</body>
</html>
"""

class TestBillScraper(unittest.TestCase):
    def test_bill_scraper_extraction(self):
        soup = BeautifulSoup(SAMPLE_HTML, "html.parser")
        out = {"_account": "30269310004301", "Remarks": "Fetched"}
        PITCBillScraper.put_grid_col_cells(out, soup)
        PITCBillScraper.put_general_pairs(out, soup)
        PITCBillScraper.put_general_text_pairs(out, soup)
        PITCBillScraper.put_meter_tables(out, soup)
        PITCBillScraper.put_meter_text_sections(out, soup)
        PITCBillScraper.put_cum_mdi(out, soup)
        PITCBillScraper.put_charges_breakdown(out, soup)
        out["Total Units"] = compute_total_units(out)

        self.assertEqual(out.get("status"), "ALIVE")
        self.assertEqual(out.get("due_date"), "23 OCT 26")
        self.assertEqual(out.get("name_address"), "Edoto Mobile Tower, Near National Club, D.I.Khan")
        self.assertEqual(out.get("kwh_present_1"), "26205.61")
        self.assertEqual(out.get("kwh_present_2"), "6754.78")
        self.assertEqual(out.get("kwh_mf_1"), "1")
        self.assertEqual(out.get("kwh_mf_2"), "1")
        self.assertEqual(out.get("kwh_units_1"), "960")
        self.assertEqual(out.get("kwh_units_2"), "261")
        self.assertEqual(out.get("Total Units"), 1221)

    def test_header_detection_helpers(self):
        # Known headers must be detected
        headers = [
            "Reference No", "Reference No.", "Ref No", "Ref#", "Reference",
            "AC No", "AC No.", "AC Number", "Account", "Account No", "Account Number",
            "Consumer ID", "Consumer No", "A/C", "A/C No", "Acc No",
            "Sr. No.", "Sr No", "Serial No", "Status", "Remarks",
            "PESCO Billing Report", "Defaulter List 2026"
        ]
        for h in headers:
            self.assertTrue(is_header_text(h), f"Expected '{h}' to be identified as header text")

        # AC-specific headers
        self.assertTrue(is_ac_header("Reference No"))
        self.assertTrue(is_ac_header("Ref No."))
        self.assertTrue(is_ac_header("AC No."))
        self.assertTrue(is_ac_header("Account Number"))
        self.assertTrue(is_ac_header("Consumer ID"))
        self.assertFalse(is_ac_header("Sr. No."))

        # SR-specific headers
        self.assertTrue(is_sr_header("Sr. No."))
        self.assertTrue(is_sr_header("Sr No"))
        self.assertTrue(is_sr_header("Serial No"))
        self.assertFalse(is_sr_header("AC No."))

        # Valid electricity account numbers must NOT be identified as headers
        valid_accounts = [
            "26111000123401",
            26111000123401,
            "01 26111 0001234 U",
            "30269310004301",
            "01261110001234",
            "1234567890",
        ]
        for acc in valid_accounts:
            self.assertFalse(is_header_text(acc), f"Expected '{acc}' to NOT be identified as header text")

    def test_update_workbook_skips_reference_no_header(self):
        wb = Workbook()
        ws = wb.active
        ws.append(["Reference No", "Name", "Tariff"])
        ws.append(["26111000123401", "Consumer One", "A-1A(01)"])
        ws.append(["26111000123402", "Consumer Two", "A-1A(01)"])

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
            tmp_path = Path(tf.name)
        wb.save(tmp_path)

        try:
            scraper = PITCBillScraper(delay=0.0)
            scraped_accounts = []

            def mock_scrape(account, **kwargs):
                scraped_accounts.append(account)
                return {
                    "_account": account,
                    "tariff": "A-1A(01)",
                    "Remarks": "Fetched",
                    "scrape_status": "Fetched",
                }

            scraper.scrape_account = mock_scrape
            scraper.update_workbook(tmp_path)

            # "Reference No" header should NEVER be scraped as an account
            self.assertNotIn("Reference No", scraped_accounts)
            self.assertEqual(scraped_accounts, ["26111000123401", "26111000123402"])

            # Verify the workbook content
            from openpyxl import load_workbook
            res_wb = load_workbook(tmp_path)
            res_ws = res_wb.active
            # Row 1 is header
            self.assertEqual(res_ws.cell(1, 1).value, "AC No.")
            # Row 2 is first account
            self.assertEqual(str(res_ws.cell(2, 1).value), "26111000123401")
            # Row 3 is second account
            self.assertEqual(str(res_ws.cell(3, 1).value), "26111000123402")
            # Total data rows should be 2 (plus header row = 3)
            self.assertEqual(res_ws.max_row, 3)
        finally:
            if tmp_path.exists():
                tmp_path.unlink()

    def test_update_workbook_skips_title_and_header_rows(self):
        wb = Workbook()
        ws = wb.active
        ws.append(["PESCO Billing Report"])  # Row 1: Title
        ws.append(["AC No.", "Sr. No."])     # Row 2: Header
        ws.append(["26111000123401", 1])     # Row 3: Data

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
            tmp_path = Path(tf.name)
        wb.save(tmp_path)

        try:
            scraper = PITCBillScraper(delay=0.0)
            scraped_accounts = []

            def mock_scrape(account, **kwargs):
                scraped_accounts.append(account)
                return {
                    "_account": account,
                    "Remarks": "Fetched",
                    "scrape_status": "Fetched",
                }

            scraper.scrape_account = mock_scrape
            scraper.update_workbook(tmp_path)

            # Neither title nor header row should be scraped
            self.assertNotIn("PESCO Billing Report", scraped_accounts)
            self.assertNotIn("AC No.", scraped_accounts)
            self.assertEqual(scraped_accounts, ["26111000123401"])

            from openpyxl import load_workbook
            res_wb = load_workbook(tmp_path)
            res_ws = res_wb.active
            self.assertEqual(res_ws.cell(1, 1).value, "AC No.")
            self.assertEqual(str(res_ws.cell(2, 1).value), "26111000123401")
            self.assertEqual(res_ws.max_row, 2)
        finally:
            if tmp_path.exists():
                tmp_path.unlink()

    def test_update_workbook_no_headers(self):
        wb = Workbook()
        ws = wb.active
        ws.append(["26111000123401", "Consumer One", "A-1A(01)"])
        ws.append(["26111000123402", "Consumer Two", "A-1A(01)"])

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
            tmp_path = Path(tf.name)
        wb.save(tmp_path)

        try:
            scraper = PITCBillScraper(delay=0.0)
            scraped_accounts = []

            def mock_scrape(account, **kwargs):
                scraped_accounts.append(account)
                return {
                    "_account": account,
                    "Remarks": "Fetched",
                    "scrape_status": "Fetched",
                }

            scraper.scrape_account = mock_scrape
            scraper.update_workbook(tmp_path)

            self.assertEqual(scraped_accounts, ["26111000123401", "26111000123402"])

            from openpyxl import load_workbook
            res_wb = load_workbook(tmp_path)
            res_ws = res_wb.active
            self.assertEqual(res_ws.cell(1, 1).value, "AC No.")
            self.assertEqual(str(res_ws.cell(2, 1).value), "26111000123401")
            self.assertEqual(str(res_ws.cell(3, 1).value), "26111000123402")
        finally:
            if tmp_path.exists():
                tmp_path.unlink()

    def test_bill_downloader_detect_rows_skips_headers(self):
        from pitcsuite.tools.bill_downloader import detect_downloader_rows

        # Case 1: Standard header row
        wb1 = Workbook()
        ws1 = wb1.active
        ws1.append(["AC No.", "Sr. No.", "Status"])
        ws1.append(["26111000123401", 1, None])
        ws1.append(["26111000123402", 2, None])

        hdr_row, ac_col, sr_col, status_col, rows = detect_downloader_rows(ws1)
        self.assertEqual(hdr_row, 1)
        self.assertEqual(ac_col, 1)
        self.assertEqual(sr_col, 2)
        self.assertEqual(status_col, 3)
        self.assertEqual(rows, [2, 3])

        # Case 2: Title row in row 1, headers in row 2
        wb2 = Workbook()
        ws2 = wb2.active
        ws2.append(["PESCO Monthly Billing Extraction"])  # Row 1: Title
        ws2.append(["Reference No", "Sr No"])              # Row 2: Header
        ws2.append(["26111000123401", 1])                  # Row 3: Data
        ws2.append(["26111000123402", 2])                  # Row 4: Data

        hdr_row, ac_col, sr_col, status_col, rows = detect_downloader_rows(ws2)
        self.assertEqual(hdr_row, 2)
        self.assertEqual(ac_col, 1)
        self.assertEqual(sr_col, 2)
        self.assertEqual(status_col, 3)
        self.assertEqual(rows, [3, 4])

        # Case 3: Column order reversed (Sr No in col 1, Reference No in col 2)
        wb3 = Workbook()
        ws3 = wb3.active
        ws3.append(["Sr. No.", "Reference No."])
        ws3.append([1, "26111000123401"])
        ws3.append([2, "26111000123402"])

        hdr_row, ac_col, sr_col, status_col, rows = detect_downloader_rows(ws3)
        self.assertEqual(hdr_row, 1)
        self.assertEqual(ac_col, 2)
        self.assertEqual(sr_col, 1)
        self.assertEqual(rows, [2, 3])

        # Case 4: No header row (pure data from row 1)
        wb4 = Workbook()
        ws4 = wb4.active
        ws4.append(["26111000123401", 1])
        ws4.append(["26111000123402", 2])

        hdr_row, ac_col, sr_col, status_col, rows = detect_downloader_rows(ws4)
        self.assertIsNone(hdr_row)
        self.assertEqual(ac_col, 1)
        self.assertEqual(sr_col, 2)
        self.assertEqual(rows, [1, 2])

if __name__ == "__main__":
    unittest.main()

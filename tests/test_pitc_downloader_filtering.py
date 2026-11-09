import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from PySide6.QtWidgets import QApplication

try:
    import pymupdf as fitz
except ImportError:
    import fitz

import unittest
import tempfile
import shutil

from pitcsuite.tools.pitc_downloader import (
    AMOUNT_DIFFERENCE_TOLERANCE,
    AmountNotFoundError,
    MonthNotFoundError,
    NoMatchError,
    amount_as_integer,
    amount_matches,
    format_month,
    get_column_boundaries,
    find_amount_rects_in_page,
    find_page_highlights,
    filter_pdf,
    filter_pdf_by_range,
)


class TestPITCDownloaderFiltering(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def _create_billing_details_pdf(self, path, rows_data):
        """
        rows_data: list of dicts:
        {
            'month': 'Sep - 2026',
            'opening_balance': 4600.0,
            'normal': 0.0,
            'detected': 0.0,
            'audit': 0.0,
            'total_payable': 4600.0,
            'total_receivable': 4600.0
        }
        """
        doc = fitz.open()
        page = doc.new_page(width=841, height=595)

        # Standard vertical lines from Crystal Reports PITC Billing Details
        v_lines = [18.5, 62.8, 119.5, 162.5, 201.5, 257.8, 306.5, 356.0, 416.5, 459.2, 514.5, 569.0, 611.0, 686.8, 749.0, 809.8]
        for x in v_lines:
            page.draw_line(fitz.Point(x, 129.7), fitz.Point(x, 517.2))

        # Title
        page.insert_text(fitz.Point(280, 50), "Billing Details", fontsize=12)

        # Headers
        headers = [
            (29.2, 145.3, "BILLING\nMONTH"),
            (219.9, 145.4, "OPENING\nBALANCE"),
            (479.6, 150.6, "NORMAL"),
            (526.6, 150.6, "DETECTED"),
            (578.8, 150.6, "AUDIT"),
            (622.7, 145.4, "TOTAL\nPAYABLE\nCONSUMMER"),
            (767.8, 145.4, "Total\nReceivable"),
        ]
        for x, y, text in headers:
            page.insert_text(fitz.Point(x, y), text, fontsize=8)

        # Insert rows
        curr_y = 175.0
        for r in rows_data:
            if "month" in r:
                page.insert_text(fitz.Point(25.0, curr_y), str(r["month"]), fontsize=8)
            if "opening_balance" in r:
                page.insert_text(fitz.Point(220.0, curr_y), f"{r['opening_balance']:,.2f}", fontsize=8)
            if "normal" in r:
                page.insert_text(fitz.Point(485.0, curr_y), f"{r['normal']:,.2f}", fontsize=8)
            if "detected" in r:
                page.insert_text(fitz.Point(535.0, curr_y), f"{r['detected']:,.2f}", fontsize=8)
            if "audit" in r:
                page.insert_text(fitz.Point(575.0, curr_y), f"{r['audit']:,.2f}", fontsize=8)
            if "total_payable" in r:
                page.insert_text(fitz.Point(635.0, curr_y), f"{r['total_payable']:,.2f}", fontsize=8)
            if "total_receivable" in r:
                page.insert_text(fitz.Point(770.0, curr_y), f"{r['total_receivable']:,.2f}", fontsize=8)
            curr_y += 20.0

        # Footer
        page.insert_text(fitz.Point(30, 535), "Not valid for any legal proceeding without any authorized signature.", fontsize=8)

        doc.save(path)
        doc.close()

    def _create_adjustment_details_pdf(self, path, rows_data):
        doc = fitz.open()
        page = doc.new_page(width=841, height=595)

        # Vertical lines
        v_lines = [20.0, 100.0, 200.0, 300.0, 400.0, 500.0, 600.0, 700.0, 800.0]
        for x in v_lines:
            page.draw_line(fitz.Point(x, 120.0), fitz.Point(x, 500.0))

        page.insert_text(fitz.Point(280, 50), "Billing Adjustment Details", fontsize=12)

        # Headers
        page.insert_text(fitz.Point(30, 140), "BILLING MONTH", fontsize=8)
        page.insert_text(fitz.Point(420, 140), "CHARGES", fontsize=8)
        page.insert_text(fitz.Point(730, 140), "TOTAL", fontsize=8)

        curr_y = 175.0
        for r in rows_data:
            if "month" in r:
                page.insert_text(fitz.Point(30.0, curr_y), str(r["month"]), fontsize=8)
            if "charges" in r:
                page.insert_text(fitz.Point(420.0, curr_y), f"{r['charges']:,.2f}", fontsize=8)
            if "total" in r:
                page.insert_text(fitz.Point(730.0, curr_y), f"{r['total']:,.2f}", fontsize=8)
            curr_y += 20.0

        page.insert_text(fitz.Point(30, 535), "Not valid for any legal proceeding without any authorized signature.", fontsize=8)

        doc.save(path)
        doc.close()

    def test_forbidden_columns_never_matched_in_billing_details(self):
        """Account 17269322131401 case: 4,600 in Opening Balance / Total Payable / Total Receivable."""
        pdf_path = os.path.join(self.temp_dir, "test_17269322131401.pdf")
        out_path = os.path.join(self.temp_dir, "out_17269322131401.pdf")
        rows = [
            {
                "month": "Sep - 2026",
                "opening_balance": 4600.0,
                "normal": 0.0,
                "detected": 0.0,
                "audit": 0.0,
                "total_payable": 4600.0,
                "total_receivable": 4600.0,
            }
        ]
        self._create_billing_details_pdf(pdf_path, rows)

        # Searching for 4600 must raise AmountNotFoundError because it's only in forbidden columns
        with self.assertRaises(AmountNotFoundError):
            filter_pdf(pdf_path, out_path, month=None, amount=4600, report_key="billing")

        # Output PDF must NOT exist
        self.assertFalse(os.path.exists(out_path))

    def test_amount_not_found_raises_and_cleans_up(self):
        """Account 19269321229425 case: amount 3914 not found in PDF."""
        pdf_path = os.path.join(self.temp_dir, "test_19269321229425.pdf")
        out_path = os.path.join(self.temp_dir, "out_19269321229425.pdf")
        rows = [
            {
                "month": "Sep - 2026",
                "opening_balance": 15423.0,
                "normal": 0.0,
                "detected": 0.0,
                "audit": 0.0,
                "total_payable": 5461.0,
                "total_receivable": 5396.0,
            }
        ]
        self._create_billing_details_pdf(pdf_path, rows)

        with self.assertRaises(AmountNotFoundError):
            filter_pdf(pdf_path, out_path, month=None, amount=3914, report_key="billing")

        self.assertFalse(os.path.exists(out_path))

    def test_normal_column_matched_and_highlighted(self):
        """Amount in NORMAL column is matched and highlighted."""
        pdf_path = os.path.join(self.temp_dir, "test_normal.pdf")
        out_path = os.path.join(self.temp_dir, "out_normal.pdf")
        rows = [
            {
                "month": "Apr - 2023",
                "opening_balance": 33702.0,
                "normal": 597.5,
                "detected": 0.0,
                "audit": 0.0,
                "total_payable": 41040.0,
                "total_receivable": 41676.0,
            }
        ]
        self._create_billing_details_pdf(pdf_path, rows)

        filter_pdf(pdf_path, out_path, month=None, amount=598, report_key="billing")
        self.assertTrue(os.path.exists(out_path))

        doc = fitz.open(out_path)
        self.assertEqual(len(doc), 1)
        page = doc[0]
        annots = list(page.annots())
        self.assertEqual(len(annots), 1)
        # Verify the highlight annot is in the NORMAL column [459.2, 514.5]
        rect = annots[0].rect
        self.assertGreaterEqual(rect.x0, 455.0)
        self.assertLessEqual(rect.x1, 520.0)
        doc.close()

    def test_audit_column_matched_and_highlighted(self):
        """Amount in AUDIT column is matched and highlighted."""
        pdf_path = os.path.join(self.temp_dir, "test_audit.pdf")
        out_path = os.path.join(self.temp_dir, "out_audit.pdf")
        rows = [
            {
                "month": "Jan - 2025",
                "opening_balance": 72497.0,
                "normal": 0.0,
                "detected": 0.0,
                "audit": 46532.0,
                "total_payable": 127398.0,
                "total_receivable": 127819.0,
            }
        ]
        self._create_billing_details_pdf(pdf_path, rows)

        filter_pdf(pdf_path, out_path, month=None, amount=46532, report_key="billing")
        self.assertTrue(os.path.exists(out_path))

        doc = fitz.open(out_path)
        self.assertEqual(len(doc), 1)
        page = doc[0]
        annots = list(page.annots())
        self.assertEqual(len(annots), 1)
        rect = annots[0].rect
        # AUDIT column is [569.0, 611.0]
        self.assertGreaterEqual(rect.x0, 565.0)
        self.assertLessEqual(rect.x1, 615.0)
        doc.close()

    def test_detected_column_matched_and_highlighted(self):
        """Amount in DETECTED column is matched and highlighted."""
        pdf_path = os.path.join(self.temp_dir, "test_detected.pdf")
        out_path = os.path.join(self.temp_dir, "out_detected.pdf")
        rows = [
            {
                "month": "Jul - 2021",
                "opening_balance": 0.0,
                "normal": 0.0,
                "detected": 21301.82,
                "audit": 0.0,
                "total_payable": 23551.82,
                "total_receivable": 23803.82,
            }
        ]
        self._create_billing_details_pdf(pdf_path, rows)

        filter_pdf(pdf_path, out_path, month=None, amount=21302, report_key="billing")
        self.assertTrue(os.path.exists(out_path))

        doc = fitz.open(out_path)
        self.assertEqual(len(doc), 1)
        page = doc[0]
        annots = list(page.annots())
        self.assertEqual(len(annots), 1)
        rect = annots[0].rect
        # DETECTED column is [514.5, 569.0]
        self.assertGreaterEqual(rect.x0, 510.0)
        self.assertLessEqual(rect.x1, 575.0)
        doc.close()

    def test_adjustment_details_total_column_matched(self):
        """In Billing Adjustment Detail, TOTAL column is matched."""
        pdf_path = os.path.join(self.temp_dir, "test_adj.pdf")
        out_path = os.path.join(self.temp_dir, "out_adj.pdf")
        rows = [
            {
                "month": "Sep - 2026",
                "charges": 500.0,
                "total": 2500.0,
            }
        ]
        self._create_adjustment_details_pdf(pdf_path, rows)

        filter_pdf(pdf_path, out_path, month=None, amount=2500, report_key="adjustment")
        self.assertTrue(os.path.exists(out_path))

        doc = fitz.open(out_path)
        self.assertEqual(len(doc), 1)
        page = doc[0]
        annots = list(page.annots())
        self.assertEqual(len(annots), 1)
        rect = annots[0].rect
        # TOTAL column is [700, 800]
        self.assertGreaterEqual(rect.x0, 690.0)
        self.assertLessEqual(rect.x1, 810.0)
        doc.close()

    def test_adjustment_details_other_column_not_matched(self):
        """In Billing Adjustment Detail, non-TOTAL columns must NOT be matched."""
        pdf_path = os.path.join(self.temp_dir, "test_adj_charges.pdf")
        out_path = os.path.join(self.temp_dir, "out_adj_charges.pdf")
        rows = [
            {
                "month": "Sep - 2026",
                "charges": 500.0,
                "total": 2500.0,
            }
        ]
        self._create_adjustment_details_pdf(pdf_path, rows)

        # Searching for 500 (which is in Charges, not TOTAL) must raise AmountNotFoundError
        with self.assertRaises(AmountNotFoundError):
            filter_pdf(pdf_path, out_path, month=None, amount=500, report_key="adjustment")

        self.assertFalse(os.path.exists(out_path))

    def test_month_and_amount_matching(self):
        """Both month and amount must match on the same row."""
        pdf_path = os.path.join(self.temp_dir, "test_both.pdf")
        out_path = os.path.join(self.temp_dir, "out_both.pdf")
        rows = [
            {
                "month": "Aug - 2026",
                "opening_balance": 1000.0,
                "normal": 500.0,
                "detected": 0.0,
                "audit": 0.0,
                "total_payable": 1500.0,
                "total_receivable": 1500.0,
            },
            {
                "month": "Sep - 2026",
                "opening_balance": 1000.0,
                "normal": 900.0,
                "detected": 0.0,
                "audit": 0.0,
                "total_payable": 1900.0,
                "total_receivable": 1900.0,
            },
        ]
        self._create_billing_details_pdf(pdf_path, rows)

        # Matching month Sep - 2026 and amount 900
        filter_pdf(pdf_path, out_path, month="Sep - 2026", amount=900, report_key="billing")
        self.assertTrue(os.path.exists(out_path))

        doc = fitz.open(out_path)
        page = doc[0]
        annots = list(page.annots())
        # Both month and normal amount are highlighted
        self.assertEqual(len(annots), 2)
        doc.close()

        # Month Sep - 2026 with amount 500 (which belongs to Aug) must NOT match
        out_path_fail = os.path.join(self.temp_dir, "out_fail.pdf")
        with self.assertRaises(AmountNotFoundError):
            filter_pdf(pdf_path, out_path_fail, month="Sep - 2026", amount=500, report_key="billing")
        self.assertFalse(os.path.exists(out_path_fail))

    def test_multipage_match_modes(self):
        """Test first, second, and all matching pages across a multi-page PDF."""
        pdf_path = os.path.join(self.temp_dir, "test_multi.pdf")
        doc = fitz.open()

        # Page 1: normal amount 1000
        p1 = doc.new_page(width=841, height=595)
        p1.insert_text(fitz.Point(300, 50), "Billing Details")
        p1.insert_text(fitz.Point(480, 150), "NORMAL")
        p1.insert_text(fitz.Point(485, 200), "1,000.00")

        # Page 2: no match
        p2 = doc.new_page(width=841, height=595)
        p2.insert_text(fitz.Point(300, 50), "Billing Details")
        p2.insert_text(fitz.Point(480, 150), "NORMAL")
        p2.insert_text(fitz.Point(485, 200), "0.00")

        # Page 3: normal amount 1000
        p3 = doc.new_page(width=841, height=595)
        p3.insert_text(fitz.Point(300, 50), "Billing Details")
        p3.insert_text(fitz.Point(480, 150), "NORMAL")
        p3.insert_text(fitz.Point(485, 200), "1,000.00")

        doc.save(pdf_path)
        doc.close()

        # Match mode first: should extract 1 page (Page 1)
        out_first = os.path.join(self.temp_dir, "out_first.pdf")
        filter_pdf(pdf_path, out_first, amount=1000, match_mode="first", report_key="billing")
        d1 = fitz.open(out_first)
        self.assertEqual(len(d1), 1)
        d1.close()

        # Match mode second: should extract 1 page (Page 3)
        out_second = os.path.join(self.temp_dir, "out_second.pdf")
        filter_pdf(pdf_path, out_second, amount=1000, match_mode="second", report_key="billing")
        d2 = fitz.open(out_second)
        self.assertEqual(len(d2), 1)
        d2.close()

        # Match mode all: should extract 2 pages (Page 1 and Page 3)
        out_all = os.path.join(self.temp_dir, "out_all.pdf")
        filter_pdf(pdf_path, out_all, amount=1000, match_mode="all", report_key="billing")
        d3 = fitz.open(out_all)
        self.assertEqual(len(d3), 2)
        d3.close()

    def test_month_only_matching(self):
        """When only month is specified, matching page is extracted and month is highlighted."""
        pdf_path = os.path.join(self.temp_dir, "test_month_only.pdf")
        rows = [
            {"month": "Jun - 2026", "normal": 0.0}
        ]
        self._create_billing_details_pdf(pdf_path, rows)

        out_path = os.path.join(self.temp_dir, "out_month_only.pdf")
        filter_pdf(pdf_path, out_path, month="Jun - 2026", amount=None, report_key="billing")
        self.assertTrue(os.path.exists(out_path))

        doc = fitz.open(out_path)
        self.assertEqual(len(doc), 1)
        annots = list(doc[0].annots())
        self.assertEqual(len(annots), 1)
        doc.close()

    def test_tolerance_matching(self):
        """Tolerance of 9 rupees: 995 matches 1000, but 980 does not."""
        pdf_path = os.path.join(self.temp_dir, "test_tol.pdf")
        rows = [{"month": "Jan - 2026", "normal": 1000.0}]
        self._create_billing_details_pdf(pdf_path, rows)

        # 995 is within tolerance of 9 rupees from 1000
        out_pass = os.path.join(self.temp_dir, "out_tol_pass.pdf")
        filter_pdf(pdf_path, out_pass, amount=995, report_key="billing")
        self.assertTrue(os.path.exists(out_pass))

        # 980 is outside tolerance
        out_fail = os.path.join(self.temp_dir, "out_tol_fail.pdf")
        with self.assertRaises(AmountNotFoundError):
            filter_pdf(pdf_path, out_fail, amount=980, report_key="billing")
        self.assertFalse(os.path.exists(out_fail))


if __name__ == "__main__":
    unittest.main()

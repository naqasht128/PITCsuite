import os
import unittest
import tempfile
from pitcsuite.tools.cp139_posting_checker import (
    parse_cp139_file, deduplicate_records, filter_records, export_cp139_to_excel
)

class TestCP139Parser(unittest.TestCase):
    def setUp(self):
        self.sample_dir = r"D:\Projects\CP139PostingChecker"
        self.file1 = os.path.join(self.sample_dir, "CP_139_POSTING_LIST_16092026_2606.txt")
        self.file2 = os.path.join(self.sample_dir, "CP_139_POSTING_LIST_24092026_2606.txt")

    def test_parse_sample_1(self):
        if not os.path.exists(self.file1):
            self.skipTest("Sample 1 file not found")
        records = parse_cp139_file(self.file1)
        self.assertEqual(len(records), 47)

        # Check first record
        r0 = records[0]
        self.assertEqual(r0["Ref#"], "30266111005700")
        self.assertEqual(r0["S/Div"], "26611")
        self.assertEqual(r0["G. Total"], 117893)
        self.assertEqual(r0["Errors"], "Posted")

        # Check a Not Pos record with error code
        not_pos = [r for r in records if r["Errors"] == "Not Pos"]
        self.assertEqual(len(not_pos), 7)
        self.assertEqual(not_pos[0]["Ref#"], "01269320212907")
        self.assertEqual(not_pos[0]["S/Div"], "26932")
        self.assertEqual(not_pos[0]["G. Total"], 1220)
        self.assertEqual(not_pos[0]["Error_Detail"], "Not Pos (G09)")

    def test_parse_sample_2(self):
        if not os.path.exists(self.file2):
            self.skipTest("Sample 2 file not found")
        records = parse_cp139_file(self.file2)
        self.assertEqual(len(records), 127)

        posted = [r for r in records if r["Errors"] == "Posted"]
        not_pos = [r for r in records if r["Errors"] == "Not Pos"]
        self.assertEqual(len(posted), 114)
        self.assertEqual(len(not_pos), 13)

        # Check all records have valid 14-digit Ref# and 5-digit S/Div
        for r in records:
            self.assertEqual(len(r["Ref#"]), 14)
            self.assertTrue(r["Ref#"].isdigit())
            self.assertEqual(len(r["S/Div"]), 5)
            self.assertEqual(r["S/Div"], r["Ref#"][2:7])
            self.assertGreater(r["G. Total"], 0)

    def test_deduplication(self):
        sample_recs = [
            {"Ref#": "01234567890123", "S/Div": "23456", "G. Total": 100, "Errors": "Not Pos", "Error_Detail": "Not Pos (G09)"},
            {"Ref#": "01234567890123", "S/Div": "23456", "G. Total": 100, "Errors": "Posted", "Error_Detail": "Posted"},
            {"Ref#": "99999999999999", "S/Div": "99999", "G. Total": 500, "Errors": "Posted", "Error_Detail": "Posted"},
        ]
        # Prefer posted
        deduped, dups = deduplicate_records(sample_recs, mode="prefer_posted")
        self.assertEqual(len(deduped), 2)
        self.assertEqual(dups, 1)
        r0 = [r for r in deduped if r["Ref#"] == "01234567890123"][0]
        self.assertEqual(r0["Errors"], "Posted")

    def test_filtering(self):
        sample_recs = [
            {"Ref#": "1", "Errors": "Posted"},
            {"Ref#": "2", "Errors": "Not Pos"},
            {"Ref#": "3", "Errors": "Posted"},
        ]
        self.assertEqual(len(filter_records(sample_recs, "posted")), 2)
        self.assertEqual(len(filter_records(sample_recs, "not_pos")), 1)
        self.assertEqual(len(filter_records(sample_recs, "all")), 3)

    def test_excel_export(self):
        if not os.path.exists(self.file1):
            self.skipTest("Sample 1 file not found")
        records = parse_cp139_file(self.file1)
        with tempfile.TemporaryDirectory() as tmpdir:
            out_file = os.path.join(tmpdir, "test_cp139_output.xlsx")
            stats = export_cp139_to_excel(records, out_file, include_source_file=True)
            self.assertTrue(os.path.exists(out_file))
            self.assertEqual(stats["total"], 47)
            self.assertEqual(stats["posted"], 40)
            self.assertEqual(stats["not_pos"], 7)
            self.assertGreater(stats["total_amount"], 0)

if __name__ == "__main__":
    unittest.main()

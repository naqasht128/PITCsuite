import os
import shutil
import tempfile
import unittest
import openpyxl

from pitcsuite.tools.integrated_tools import FileMoverWorker


class TestFileMoverWorker(unittest.TestCase):
    def setUp(self):
        self.test_dir = tempfile.mkdtemp()
        self.source_dir = os.path.join(self.test_dir, "source")
        self.dest_dir = os.path.join(self.test_dir, "dest")
        self.excel_path = os.path.join(self.test_dir, "input.xlsx")
        os.makedirs(self.source_dir, exist_ok=True)
        os.makedirs(self.dest_dir, exist_ok=True)

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_file_mover_with_header(self):
        # Create dummy source files
        with open(os.path.join(self.source_dir, "101.pdf"), "w") as f:
            f.write("content 101")
        with open(os.path.join(self.source_dir, "103.pdf"), "w") as f:
            f.write("content 103")

        # Create input Excel with header
        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = "Sr No"
        ws["A2"] = 101
        ws["A3"] = 102
        ws["A4"] = 103
        wb.save(self.excel_path)
        wb.close()

        # Run worker directly
        worker = FileMoverWorker(self.excel_path, self.source_dir, self.dest_dir)
        results = {}
        worker.done.connect(lambda moved, missing: results.update({"moved": moved, "missing": missing}))
        worker.failed.connect(lambda err: results.update({"error": err}))
        worker.run()

        self.assertNotIn("error", results)
        self.assertEqual(results.get("moved"), 2)
        self.assertEqual(results.get("missing"), ["102"])

        # Check physical files
        self.assertTrue(os.path.isfile(os.path.join(self.dest_dir, "101.pdf")))
        self.assertTrue(os.path.isfile(os.path.join(self.dest_dir, "103.pdf")))
        self.assertFalse(os.path.isfile(os.path.join(self.source_dir, "101.pdf")))
        self.assertFalse(os.path.isfile(os.path.join(self.source_dir, "103.pdf")))

        # Check Excel updates in Column B
        wb_check = openpyxl.load_workbook(self.excel_path)
        ws_check = wb_check.active
        self.assertEqual(ws_check["B1"].value, "Status")
        self.assertEqual(ws_check["B2"].value, "Moved")
        self.assertEqual(ws_check["B3"].value, "Not Found")
        self.assertEqual(ws_check["B4"].value, "Moved")
        wb_check.close()

    def test_file_mover_without_header(self):
        with open(os.path.join(self.source_dir, "5001.pdf"), "w") as f:
            f.write("content 5001")

        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = "5001"
        ws["A2"] = "5002"
        wb.save(self.excel_path)
        wb.close()

        worker = FileMoverWorker(self.excel_path, self.source_dir, self.dest_dir)
        results = {}
        worker.done.connect(lambda moved, missing: results.update({"moved": moved, "missing": missing}))
        worker.failed.connect(lambda err: results.update({"error": err}))
        worker.run()

        self.assertNotIn("error", results)
        self.assertEqual(results.get("moved"), 1)
        self.assertEqual(results.get("missing"), ["5002"])

        wb_check = openpyxl.load_workbook(self.excel_path)
        ws_check = wb_check.active
        self.assertEqual(ws_check["B1"].value, "Moved")
        self.assertEqual(ws_check["B2"].value, "Not Found")
        wb_check.close()

    def test_file_mover_with_extensions_and_duplicates(self):
        with open(os.path.join(self.source_dir, "report.pdf"), "w") as f:
            f.write("report")

        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = "File Name"
        ws["B1"] = "Status"
        ws["A2"] = "report.pdf"
        ws["A3"] = "missing_file.pdf"
        ws["A4"] = "report.pdf"  # duplicate
        wb.save(self.excel_path)
        wb.close()

        worker = FileMoverWorker(self.excel_path, self.source_dir, self.dest_dir)
        results = {}
        worker.done.connect(lambda moved, missing: results.update({"moved": moved, "missing": missing}))
        worker.failed.connect(lambda err: results.update({"error": err}))
        worker.run()

        self.assertNotIn("error", results)
        self.assertEqual(results.get("moved"), 1)
        self.assertEqual(results.get("missing"), ["missing_file.pdf"])

        wb_check = openpyxl.load_workbook(self.excel_path)
        ws_check = wb_check.active
        self.assertEqual(ws_check["B1"].value, "Status")
        self.assertEqual(ws_check["B2"].value, "Moved")
        self.assertEqual(ws_check["B3"].value, "Not Found")
        self.assertEqual(ws_check["B4"].value, "Moved")
        wb_check.close()

    def test_file_mover_float_and_case_insensitive(self):
        with open(os.path.join(self.source_dir, "Document_ABC.PDF"), "w") as f:
            f.write("abc")
        with open(os.path.join(self.source_dir, "2050.pdf"), "w") as f:
            f.write("2050")

        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = "Files"
        ws["A2"] = 2050.0  # float representation of int
        ws["A3"] = "document_abc"  # lowercase match against Document_ABC.PDF
        ws["A4"] = None  # blank line
        ws["A5"] = "non_existent"
        wb.save(self.excel_path)
        wb.close()

        worker = FileMoverWorker(self.excel_path, self.source_dir, self.dest_dir)
        results = {}
        worker.done.connect(lambda moved, missing: results.update({"moved": moved, "missing": missing}))
        worker.failed.connect(lambda err: results.update({"error": err}))
        worker.run()

        self.assertNotIn("error", results)
        self.assertEqual(results.get("moved"), 2)
        self.assertEqual(results.get("missing"), ["non_existent"])

        wb_check = openpyxl.load_workbook(self.excel_path)
        ws_check = wb_check.active
        self.assertEqual(ws_check["B1"].value, "Status")
        self.assertEqual(ws_check["B2"].value, "Moved")
        self.assertEqual(ws_check["B3"].value, "Moved")
        self.assertIsNone(ws_check["B4"].value)
        self.assertEqual(ws_check["B5"].value, "Not Found")
        wb_check.close()


if __name__ == "__main__":
    unittest.main()

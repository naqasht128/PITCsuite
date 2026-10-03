import os
import unittest
import tempfile
import openpyxl

from pitcsuite.tools.assessment_realization import (
    parse_amount,
    parse_assessment_realization_file,
    deduplicate_records,
    filter_records,
    export_assessment_realization_to_excel
)

class TestAssessmentRealization(unittest.TestCase):
    def setUp(self):
        self.sample_txt = r"D:\Projects\Assessment Realization.txt"
        self.ref_xlsx = r"C:\Users\Naqash\Desktop\Assessment Realization Output.xlsx"

    def test_parse_amount(self):
        self.assertEqual(parse_amount("486145.00"), 486145)
        self.assertEqual(parse_amount("0.00"), 0)
        self.assertEqual(parse_amount("64.15"), 64.15)
        self.assertEqual(parse_amount("1,094.00"), 1094)
        self.assertEqual(parse_amount("100.00-"), -100)
        self.assertEqual(parse_amount("(50.25)"), -50.25)
        self.assertEqual(parse_amount(""), 0)
        self.assertEqual(parse_amount(None), 0)

    def test_parse_sample_file(self):
        if not os.path.exists(self.sample_txt):
            self.skipTest(f"Sample file {self.sample_txt} not found.")

        records = parse_assessment_realization_file(self.sample_txt)
        self.assertEqual(len(records), 17)

        # Record 1 (first entry in user's prompt)
        r0 = records[0]
        self.assertEqual(r0["Sdiv"], "26612")
        self.assertEqual(r0["Reference No"], "19266120816300")
        self.assertIsNone(r0["Status"])
        self.assertEqual(r0["Name & Address"], "ORAKZAI SCOUTS MIR ALI KOHAT ROAD BANNU MIRALI KOHAT ROAD BANNU")
        self.assertEqual(r0["Tariff"], "66")
        self.assertEqual(r0["Department Code"], 11022)
        self.assertEqual(r0["Arrear Age"], 190)
        self.assertEqual(r0["OP. Balance"], 486145)
        self.assertEqual(r0["Payments Normal"], 0)
        self.assertEqual(r0["Adjustor"], 0)
        self.assertEqual(r0["Write Off"], 0)
        self.assertEqual(r0["Net Arrear"], 487425)

        # Record 2 (second entry in user's prompt)
        r1 = records[1]
        self.assertEqual(r1["Sdiv"], "26612")
        self.assertEqual(r1["Reference No"], "19266120829701")
        self.assertEqual(r1["Status"], "P-DISC")
        self.assertEqual(r1["Name & Address"], "COMMANDANT TOCHI SCOUTS NEW ADDA BANNU")
        self.assertEqual(r1["Tariff"], "66")
        self.assertEqual(r1["Department Code"], 11022)
        self.assertEqual(r1["Arrear Age"], 177)
        self.assertEqual(r1["OP. Balance"], 58751)
        self.assertEqual(r1["Payments Normal"], 0)
        self.assertEqual(r1["Adjustor"], 0)
        self.assertEqual(r1["Write Off"], 0)
        self.assertEqual(r1["Net Arrear"], 58751)

        # Check all 17 records have valid reference numbers
        for r in records:
            self.assertEqual(len(r["Reference No"]), 14)
            self.assertTrue(r["Reference No"].isdigit())
            self.assertEqual(len(r["Sdiv"]), 5)
            self.assertEqual(r["Sdiv"], r["Reference No"][2:7])

        # Check status distribution
        p_disc = [r for r in records if r["Status"] == "P-DISC"]
        active = [r for r in records if not r["Status"]]
        self.assertEqual(len(p_disc), 5)
        self.assertEqual(len(active), 12)

    def test_exact_match_with_user_template(self):
        if not os.path.exists(self.sample_txt) or not os.path.exists(self.ref_xlsx):
            self.skipTest("Reference or sample file not available for exact match test.")

        records = parse_assessment_realization_file(self.sample_txt)
        wb_ref = openpyxl.load_workbook(self.ref_xlsx)
        ws_ref = wb_ref.active

        ref_cols = [ws_ref.cell(1, c).value for c in range(1, 13)]
        expected_cols = [
            "Sdiv", "Reference No", "Status", "Name & Address", "Tariff",
            "Department Code", "Arrear Age", "OP. Balance", "Payments Normal",
            "Adjustor", "Write Off", "Net Arrear"
        ]
        self.assertEqual(ref_cols, expected_cols)

        # Compare row 2
        ref_row2 = [ws_ref.cell(2, c).value for c in range(1, 13)]
        parsed_row2 = [records[0][k] for k in expected_cols]
        self.assertEqual(ref_row2, parsed_row2)

        # Compare row 3
        ref_row3 = [ws_ref.cell(3, c).value for c in range(1, 13)]
        parsed_row3 = [records[1][k] for k in expected_cols]
        self.assertEqual(ref_row3, parsed_row3)

    def test_export_excel(self):
        sample_records = [
            {
                "Sdiv": "26612",
                "Reference No": "19266120816300",
                "Status": None,
                "Name & Address": "TEST CONSUMER 1",
                "Tariff": "66",
                "Department Code": 11022,
                "Arrear Age": 10,
                "OP. Balance": 1000,
                "Payments Normal": 100,
                "Adjustor": 0,
                "Write Off": 0,
                "Net Arrear": 900,
                "Source_File": "test.txt",
            },
            {
                "Sdiv": "26612",
                "Reference No": "19266120829701",
                "Status": "P-DISC",
                "Name & Address": "TEST CONSUMER 2",
                "Tariff": "66",
                "Department Code": 11022,
                "Arrear Age": 20,
                "OP. Balance": 2000,
                "Payments Normal": 0,
                "Adjustor": 50,
                "Write Off": 0,
                "Net Arrear": 1950,
                "Source_File": "test.txt",
            }
        ]

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp_path = tmp.name

        try:
            stats = export_assessment_realization_to_excel(sample_records, tmp_path)
            self.assertEqual(stats["total"], 2)
            self.assertEqual(stats["p_disc"], 1)
            self.assertEqual(stats["active"], 1)
            self.assertEqual(stats["total_op_balance"], 3000)

            wb = openpyxl.load_workbook(tmp_path)
            ws = wb.active
            self.assertEqual(ws.title, "Assessment_Realization")
            # Headers
            self.assertEqual(ws.cell(1, 1).value, "Sdiv")
            self.assertEqual(ws.cell(1, 2).value, "Reference No")
            self.assertEqual(ws.cell(1, 12).value, "Net Arrear")

            # Check rows
            self.assertEqual(ws.cell(2, 1).value, "26612")
            self.assertEqual(ws.cell(2, 2).value, "19266120816300")
            self.assertIsNone(ws.cell(2, 3).value)
            self.assertEqual(ws.cell(3, 3).value, "P-DISC")

            # Check total row formulas
            self.assertEqual(ws.cell(4, 1).value, "TOTAL")
            self.assertEqual(ws.cell(4, 8).value, "=SUM(H2:H3)")
            self.assertEqual(ws.cell(4, 12).value, "=SUM(L2:L3)")
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    def test_deduplicate_and_filter(self):
        records = [
            {"Reference No": "01266120000001", "Status": None},
            {"Reference No": "01266120000001", "Status": "P-DISC"},
            {"Reference No": "01266120000002", "Status": "P-DISC"},
        ]
        deduped, count = deduplicate_records(records)
        self.assertEqual(len(deduped), 2)
        self.assertEqual(count, 1)

        p_disc_only = filter_records(records, "p_disc")
        self.assertEqual(len(p_disc_only), 2)

        active_only = filter_records(records, "active")
        self.assertEqual(len(active_only), 1)

if __name__ == "__main__":
    unittest.main()

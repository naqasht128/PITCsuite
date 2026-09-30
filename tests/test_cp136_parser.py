import unittest
import os
from pitcsuite.tools.cp136_posting_checker import parse_single_file, deduplicate_records

class TestCP136Parser(unittest.TestCase):
    def test_cp136_sample_pdf(self):
        pdf_path = os.path.expanduser(r"~\Desktop\CP_136_Change_PROOFLIST_22092026_2606 (1).pdf")
        if not os.path.exists(pdf_path):
            self.skipTest("Sample PDF not found on desktop")

        records = parse_single_file(pdf_path)
        self.assertEqual(len(records), 209)

        # First record check
        r0 = records[0]
        self.assertEqual(r0['Ref No'], '02269340118400')
        self.assertEqual(r0['Customer ID'], '1262659606')
        self.assertEqual(r0['Status'], 'NOT POSTED')
        self.assertEqual(r0['Report Type'], 'CP-136')
        self.assertEqual(r0.get('FEEDER CODE'), '134904')
        self.assertIn('A11', r0.get('ERRORS CODES', ''))
        self.assertIn('TR02', r0.get('ERRORS CODES', ''))
        self.assertNotIn('S-POINTS', r0)
        self.assertNotIn('EMAIL ADDRESS', r0)
        self.assertNotIn('NET MTR CD', r0)
        self.assertNotIn('METER READING', r0)

        # Deduplication
        final_recs, dups = deduplicate_records(records, mode='prefer_posted')
        self.assertEqual(len(final_recs), 208)
        self.assertEqual(dups, 1)

    def test_cp137_sample_pdf(self):
        pdf_path = os.path.expanduser(r"~\Desktop\backlog + dg capacity updation\CP_137_CHANGE_POSTING_03092026_2606.pdf")
        if not os.path.exists(pdf_path):
            self.skipTest("CP-137 sample PDF not found on desktop")

        records = parse_single_file(pdf_path)
        self.assertEqual(len(records), 10)

        # Record 0 check
        r0 = records[0]
        self.assertEqual(r0['Ref No'], '30269110854402')
        self.assertEqual(r0['Customer ID'], '1262906212')
        self.assertEqual(r0['Status'], 'POSTED')
        self.assertEqual(r0['Report Type'], 'CP-137')
        self.assertEqual(r0.get('NET MTR CD'), '01')
        self.assertEqual(r0.get('NET MTR CAP'), '20.70')
        self.assertEqual(r0.get('NET MTR CONNECTION DATE'), '6/26/2024 12:00:00AM')
        self.assertEqual(r0.get('NET MTR EXPIRY DATE'), '6/25/2031 12:00:00AM')
        self.assertNotIn('EXPIRY DATE', r0)
        self.assertNotIn('CONSUMER NAME', r0)
        self.assertNotIn('ERRORS CODES', r0)

        # Record 1 check with LICENSE#
        r1 = records[1]
        self.assertEqual(r1['Ref No'], '30269130001101')
        self.assertEqual(r1['Customer ID'], '2260052760')
        self.assertEqual(r1.get('LICENSE#'), 'NMC/17020/2025')
        self.assertEqual(r1.get('NET MTR CAP'), '244.80')

if __name__ == '__main__':
    unittest.main()


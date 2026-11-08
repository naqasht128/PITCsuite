import os
import unittest
from pitcsuite.license_manager import (
    normalize_key,
    get_device_id,
    get_candidate_device_ids,
    get_all_storage_paths,
    format_plan_type,
    format_expiry_date,
)

class TestLicenseManager(unittest.TestCase):
    def test_normalize_key(self):
        raw = "abcd-1234-cdef-5678"
        self.assertEqual(normalize_key(raw), "ABCD1234CDEF5678")
        self.assertEqual(normalize_key("  abcd1234cdef5678  "), "ABCD1234CDEF5678")

    def test_device_id(self):
        dev_id = get_device_id()
        self.assertIsInstance(dev_id, str)
        self.assertGreater(len(dev_id), 0)

        candidates = get_candidate_device_ids()
        self.assertIsInstance(candidates, list)
        self.assertGreater(len(candidates), 0)
        self.assertIn(dev_id, candidates)

    def test_storage_paths(self):
        paths = get_all_storage_paths()
        self.assertGreater(len(paths), 0)
        has_appdata = any("AppData" in p or "pitcsuite" in p.lower() for p in paths)
        self.assertTrue(has_appdata)

    def test_formatting(self):
        self.assertEqual(format_plan_type("annual"), "Annual")
        self.assertEqual(format_plan_type("lifetime"), "Lifetime")

if __name__ == '__main__':
    unittest.main()

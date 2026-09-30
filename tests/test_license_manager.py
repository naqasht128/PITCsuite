import os
import unittest
from pitcsuite.license_manager import (
    generate_expected_key,
    normalize_key,
    is_key_valid_for_machine,
    get_all_machine_hwids,
    get_license_search_paths,
    load_saved_license_keys,
    save_license_key,
    SECRET_SALT
)

class TestLicenseManager(unittest.TestCase):
    def test_key_generation(self):
        hwid = "TEST-HWID-1234"
        key = generate_expected_key(hwid)
        self.assertRegex(key, r'^[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}$')
        # Deterministic check
        self.assertEqual(key, generate_expected_key(hwid))

    def test_normalize_key(self):
        raw = "abcd-1234-CDEF-5678"
        self.assertEqual(normalize_key(raw), "ABCD1234CDEF5678")
        self.assertEqual(normalize_key("  abcd1234cdef5678  "), "ABCD1234CDEF5678")

    def test_machine_hwids_and_backward_compatibility(self):
        hwids = get_all_machine_hwids()
        self.assertIsInstance(hwids, list)
        self.assertGreater(len(hwids), 0)
        # UNKNOWN-HWID-001 must always be recognized for backward compatibility
        self.assertIn("UNKNOWN-HWID-001", hwids)

        legacy_key = generate_expected_key("UNKNOWN-HWID-001")
        is_valid, matched_hwid, exp = is_key_valid_for_machine(legacy_key)
        self.assertTrue(is_valid)
        self.assertEqual(matched_hwid, "UNKNOWN-HWID-001")

    def test_invalid_key(self):
        is_valid, _, _ = is_key_valid_for_machine("0000-0000-0000-0000")
        self.assertFalse(is_valid)
        is_valid, _, _ = is_key_valid_for_machine("invalid-short-key")
        self.assertFalse(is_valid)

    def test_license_search_paths(self):
        paths = get_license_search_paths()
        self.assertGreater(len(paths), 2)
        # Check that roaming appdata and programdata paths are present
        has_appdata = any("AppData" in p for p in paths)
        has_programdata = any("ProgramData" in p for p in paths)
        self.assertTrue(has_appdata)
        self.assertTrue(has_programdata)

if __name__ == '__main__':
    unittest.main()

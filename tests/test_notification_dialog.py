import os
import unittest
from PySide6.QtWidgets import QApplication, QWidget, QTextEdit, QPushButton
from PySide6.QtCore import Qt, QEvent
from PySide6.QtGui import QKeyEvent

from pitcsuite.ui_helpers import NotificationDialog, notify
from pitcsuite.tools.file_renamer import FileRenamerPanel


class TestNotificationDialog(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_short_message_no_text_edit(self):
        parent = QWidget()
        dlg = NotificationDialog(parent, "Test Title", "Simple short message.")
        self.assertEqual(dlg.windowTitle(), "Test Title")
        # No QTextEdit needed for short single-line messages
        text_edits = dlg.findChildren(QTextEdit)
        self.assertEqual(len(text_edits), 0)
        # OK button exists
        ok_btn = dlg.findChild(QPushButton, "ok_btn")
        self.assertIsNotNone(ok_btn)
        self.assertTrue(ok_btn.isDefault())
        dlg.close()

    def test_long_message_auto_splits_and_constrains_height(self):
        parent = QWidget()
        long_issues = [f"Row {i}: source not found - 05269310367500" for i in range(2, 50)]
        msg = "Processed 1 file(s).\n\nSkipped / issues:\n" + "\n".join(long_issues)

        dlg = NotificationDialog(parent, "File Renamer", msg, folder_path=r"C:\DummyFolder")
        dlg.show()

        # Check QTextEdit was created for the details
        text_edits = dlg.findChildren(QTextEdit)
        self.assertEqual(len(text_edits), 1)
        self.assertIn("Row 2: source not found", text_edits[0].toPlainText())
        self.assertIn("Row 49: source not found", text_edits[0].toPlainText())

        # Check height constraint: must never be taller than 520px
        hint = dlg.sizeHint()
        self.assertLessEqual(hint.height(), 520)

        # OK button exists and is visible
        ok_btn = dlg.findChild(QPushButton, "ok_btn")
        self.assertIsNotNone(ok_btn)

        # Folder button exists
        buttons = dlg.findChildren(QPushButton)
        btn_texts = [b.text() for b in buttons]
        self.assertIn("Open Folder", btn_texts)

        dlg.close()

    def test_escape_key_closes_dialog(self):
        parent = QWidget()
        dlg = NotificationDialog(parent, "Escape Test", "Testing escape key.")
        dlg.show()

        finished_results = []
        dlg.finished.connect(lambda r: finished_results.append(r))

        ev = QKeyEvent(QEvent.KeyPress, Qt.Key_Escape, Qt.NoModifier)
        QApplication.sendEvent(dlg, ev)

        self.assertEqual(len(finished_results), 1)

    def test_enter_key_closes_dialog(self):
        parent = QWidget()
        dlg = NotificationDialog(parent, "Enter Test", "Testing enter key.")
        dlg.show()

        finished_results = []
        dlg.finished.connect(lambda r: finished_results.append(r))

        ev = QKeyEvent(QEvent.KeyPress, Qt.Key_Return, Qt.NoModifier)
        QApplication.sendEvent(dlg, ev)

        self.assertEqual(len(finished_results), 1)

    def test_notify_helper_active_list(self):
        parent = QWidget()
        dlg = notify(parent, "Helper Test", "Testing notify helper.")
        self.assertIn(dlg, parent._active_notifications)
        dlg.accept()
        self.assertNotIn(dlg, parent._active_notifications)

    def test_file_renamer_done_populates_log_and_dialog(self):
        panel = FileRenamerPanel()
        issues = [f"Row {i}: source not found - 001" for i in range(2, 10)]
        panel._done(1, len(issues), issues)

        # Verify log_box contains the issues
        log_text = panel.log_box.toPlainText()
        self.assertIn("⚠️ Row 2: source not found - 001", log_text)
        self.assertIn("⚠️ Row 9: source not found - 001", log_text)

        # Verify active notification dialog was opened with details
        active = getattr(panel, "_active_notifications", [])
        self.assertEqual(len(active), 1)
        active_dlg = active[0]
        text_edits = active_dlg.findChildren(QTextEdit)
        self.assertEqual(len(text_edits), 1)
        self.assertIn("Skipped / issues:", text_edits[0].toPlainText())
        active_dlg.close()


if __name__ == "__main__":
    unittest.main()

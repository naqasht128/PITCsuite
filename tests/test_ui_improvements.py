import sys
import unittest
from pathlib import Path
from PySide6.QtWidgets import QApplication, QPushButton
from PySide6.QtCore import QSettings

app = QApplication.instance() or QApplication(sys.argv)

from pitcsuite.main_window import MainWindow, TOOLS, AboutDialog
from pitcsuite.tools.cp139_posting_checker import CP139PostingCheckerPanel
from pitcsuite.tools.settings import SettingsPanel
from pitcsuite.tools.bill_downloader import BillSoftPanel
from pitcsuite.tools.assessment_realization import AssessmentRealizationPanel
from pitcsuite.tools.payment_extract_88l import PaymentExtract88LPanel
from pitcsuite.tools.pitc_downloader import PITCPanel
from pitcsuite.tools.integrated_tools import CP136PostingCheckerPanel, CP22TCheckerPanel
from pitcsuite.ui_helpers import THEMES, apply_theme


class TestUIImprovements(unittest.TestCase):
    def setUp(self):
        self.settings = QSettings("TNDevelopers", "PITCSuite_Test")

    def test_tool_names_and_order(self):
        tool_names = [t[0] for t in TOOLS]
        # 4. Payment from 88L
        self.assertIn("Payment from 88L", tool_names)
        self.assertNotIn("Payment Extractor from 88L", tool_names)

        # 6. CP 139 Posting/PDF
        self.assertIn("CP 139 Posting/PDF", tool_names)
        self.assertNotIn("CP 139 Posting Check/PDF", tool_names)

        # 5. CP 136/137 Posting moved after CP 139 Posting/PDF
        self.assertIn("CP 136/137 Posting", tool_names)
        self.assertNotIn("CP 136 & 137 Posting Checker", tool_names)
        idx_139 = tool_names.index("CP 139 Posting/PDF")
        idx_136 = tool_names.index("CP 136/137 Posting")
        self.assertEqual(idx_136, idx_139 + 1)

        # 7. E-Bills Downloader
        self.assertIn("E-Bills Downloader", tool_names)
        self.assertNotIn("E/Bills Downloader", tool_names)

        # 8. PITC Reports Downloader
        self.assertIn("PITC Reports Downloader", tool_names)
        self.assertNotIn("PITC Report Downloader", tool_names)

        # 9. CP 22T Posting
        self.assertIn("CP 22T Posting", tool_names)
        self.assertNotIn("CP 22T Posting Checker", tool_names)

        # 10. Assessment Realization
        self.assertIn("Assessment Realization", tool_names)
        self.assertNotIn("Govt Assessment Realization", tool_names)

    def test_cp139_checkboxes_behavior(self):
        panel = CP139PostingCheckerPanel()
        
        # 1. Single PDF & Stamp Sr. No. may not be pre-checked
        self.assertFalse(panel.single_pdf_chk.isChecked())
        self.assertFalse(panel.stamp_para_chk.isChecked())
        self.assertFalse(panel.separate_pdf_chk.isChecked())
        self.assertFalse(panel.make_pdf_chk.isChecked())

        # 2. When user checks "Make PDF", separate PDF is auto checked with it,
        # single PDF is unchecked, both are enabled.
        panel.make_pdf_chk.setChecked(True)
        self.assertTrue(panel.separate_pdf_chk.isChecked())
        self.assertFalse(panel.single_pdf_chk.isChecked())
        self.assertTrue(panel.separate_pdf_chk.isEnabled())
        self.assertTrue(panel.single_pdf_chk.isEnabled())
        self.assertTrue(panel.stamp_para_chk.isEnabled())
        self.assertFalse(panel.stamp_para_chk.isChecked())  # Only used when user explicitly checks it

        # 3. If user clicks "Single PDF", it unchecks "Separate PDF"
        panel.single_pdf_chk.setChecked(True)
        self.assertTrue(panel.single_pdf_chk.isChecked())
        self.assertFalse(panel.separate_pdf_chk.isChecked())

        # If user clicks "Separate PDF", it unchecks "Single PDF"
        panel.separate_pdf_chk.setChecked(True)
        self.assertTrue(panel.separate_pdf_chk.isChecked())
        self.assertFalse(panel.single_pdf_chk.isChecked())

        # 4. Stamp Sr. No. can be checked independently
        panel.stamp_para_chk.setChecked(True)
        self.assertTrue(panel.stamp_para_chk.isChecked())

        # When Make PDF is unchecked, all disabled and unchecked
        panel.make_pdf_chk.setChecked(False)
        self.assertFalse(panel.single_pdf_chk.isChecked())
        self.assertFalse(panel.separate_pdf_chk.isChecked())
        self.assertFalse(panel.stamp_para_chk.isChecked())
        self.assertFalse(panel.single_pdf_chk.isEnabled())
        self.assertFalse(panel.separate_pdf_chk.isEnabled())
        self.assertFalse(panel.stamp_para_chk.isEnabled())

    def test_pitc_downloader_template_btn(self):
        panel = PITCPanel()
        # Find template button
        tpl_btn = None
        for b in panel.findChildren(QPushButton):
            if "Template" in b.text():
                tpl_btn = b
                break
        self.assertIsNotNone(tpl_btn)
        self.assertEqual(tpl_btn.text(), "Download Template")
        # Ensure no garish green inline stylesheet
        self.assertEqual(tpl_btn.styleSheet(), "")

    def test_no_emojis_in_buttons(self):
        # E-Bills Downloader pause button
        ebills = BillSoftPanel()
        self.assertEqual(ebills.btn_pause.text(), "PAUSE")
        ebills.pause_resume()
        self.assertEqual(ebills.btn_pause.text(), "RESUME")

        # Assessment Realization buttons
        ar = AssessmentRealizationPanel()
        self.assertEqual(ar.add_file_btn.text(), "Add File")
        self.assertEqual(ar.browse_excel_btn.text(), "Browse...")
        self.assertEqual(ar.open_excel_btn.text(), "Open Excel")
        self.assertEqual(ar.open_folder_btn.text(), "Open Folder")

        # CP139 buttons
        cp139 = CP139PostingCheckerPanel()
        self.assertEqual(cp139.browse_excel_btn.text(), "Browse Excel...")
        self.assertEqual(cp139.download_tmpl_btn.text(), "Template")

        # Payment 88L
        pay = PaymentExtract88LPanel()
        template_btns = [b for b in pay.findChildren(QPushButton) if "Template" in b.text()]
        self.assertTrue(any(b.text() == "Download Template" for b in template_btns))

    def test_sidebar_default_expanded_and_no_auto_collapse_on_save(self):
        # Clear settings to test clean first-time launch
        settings = QSettings("TNDevelopers", "PITCSuite")
        settings.remove("sidebar_style")

        win = MainWindow()
        # 1. By default, left side activity tabs show on first launch (not collapsed)
        self.assertFalse(win.nav_collapsed)

        # 12. In settings panel, saving personalization keeps left side pane as is
        settings_panel = None
        for p in win.panels:
            if isinstance(p, SettingsPanel):
                settings_panel = p
                break
        self.assertIsNotNone(settings_panel)

        # When settings panel is displayed, sidebar_style reflects the current state
        settings_panel.load_all()
        self.assertEqual(settings_panel.sidebar_style_cb.currentData(), "expanded")

        # Saving personalization must not collapse the sidebar
        settings_panel.save_personalization()
        self.assertFalse(win.nav_collapsed)

    def test_no_activation_badge_in_top_header(self):
        win = MainWindow()
        # Item 17: Activation status in every window at top is not needed
        self.assertFalse(hasattr(win, "license_badge"))

    def test_update_and_info_svg_icons(self):
        win = MainWindow()
        # about_sidebar_btn has icon and valid tooltip; update button is in AboutDialog
        self.assertFalse(win.about_sidebar_btn.icon().isNull())
        from pitcsuite.main_window import AboutDialog
        about = AboutDialog(win)
        self.assertTrue(hasattr(about, "update_action_btn"))

        # Check SVG files exist on disk
        assets_dir = Path(__file__).parent.parent / "pitcsuite" / "assets"
        self.assertTrue((assets_dir / "update.svg").exists())
        self.assertTrue((assets_dir / "info.svg").exists())

        # Info SVG must have a circle element
        info_svg_content = (assets_dir / "info.svg").read_text(encoding="utf-8")
        self.assertIn("<circle", info_svg_content)

    def test_light_theme_main_heading_is_black(self):
        win = MainWindow()
        # Apply light theme
        win._on_theme_updated("light", "blue", "compact", "compact", "inter")
        self.assertIn("#000000", win.nav_brand_lbl.styleSheet())
        self.assertIn("#000000", win.header_title.styleSheet())

    def test_status_label_color_and_subtitle_and_update_state(self):
        from pitcsuite.ui_helpers import get_status_color
        self.assertEqual(get_status_color("light"), "#000000")
        self.assertEqual(get_status_color("vesper"), "#ffffff")
        self.assertEqual(get_status_color("dark"), "#ffffff")

        win = MainWindow()
        # Subtitle check
        self.assertEqual(win.header_subtitle.text(), "Automated PITC Reports, OTP login, Filtering")

        # Sidebar About button default
        self.assertEqual(win.about_sidebar_btn.text(), "About")

        # When update arrives: emerald green and "Update Available"
        win.notify_update_available(latest_tag="v3.1.0")
        self.assertEqual(win.about_sidebar_btn.text(), "Update Available")
        self.assertIn("#10b981", win.about_sidebar_btn.styleSheet())

        # Test AboutDialog buttons
        from pitcsuite.main_window import AboutDialog
        dlg = AboutDialog(win)
        self.assertEqual(dlg.windowTitle(), "About")
        self.assertTrue(hasattr(dlg, "update_action_btn"))
        self.assertIn("Update Available", dlg.update_action_btn.text())

    def test_dark_roundish_app_icon(self):
        from pitcsuite.icons import app_icon
        icon = app_icon()
        self.assertFalse(icon.isNull())
        assets_dir = Path(__file__).parent.parent / "pitcsuite" / "assets"
        svg_content = (assets_dir / "pitcsuite_icon.svg").read_text(encoding="utf-8")
        self.assertIn('rx="54"', svg_content)
        self.assertIn("#0f172a", svg_content)

    def test_output_settings_groupbox_title(self):
        from pitcsuite.tools.integrated_tools import CP136PostingCheckerPanel
        from PySide6.QtWidgets import QGroupBox
        cp136 = CP136PostingCheckerPanel()
        titles = [g.title() for g in cp136.findChildren(QGroupBox)]
        self.assertIn("Output Settings", titles)
        self.assertNotIn("Output & Deduplication Settings", titles)

    def test_accent_buttons_light_mode_background(self):
        from pitcsuite.tools.settings import SettingsPanel
        panel = SettingsPanel()
        # Trigger update to light mode
        panel._update_accent_buttons_style("light")
        for key, btn in panel.accent_btns.items():
            style = btn.styleSheet()
            self.assertIn("background-color: #ffffff;", style)
            self.assertNotIn("#252526", style)

        # Trigger update back to dark/vesper mode
        panel._update_accent_buttons_style("vesper")
        for key, btn in panel.accent_btns.items():
            style = btn.styleSheet()
            self.assertNotIn("background-color: #ffffff;", style)

    def test_no_separator_between_settings_and_about(self):
        win = MainWindow()
        nav_layout = win.nav.layout()
        
        # Collect widgets in nav_layout
        widgets = []
        for i in range(nav_layout.count()):
            item = nav_layout.itemAt(i)
            if item.widget():
                widgets.append(item.widget())

        # Check that about_sidebar_btn is directly after the last tool button (Settings)
        self.assertIn(win.about_sidebar_btn, widgets)
        about_idx = widgets.index(win.about_sidebar_btn)
        last_tool_btn = win.nav_btns[-1]
        last_tool_idx = widgets.index(last_tool_btn)

        self.assertEqual(about_idx, last_tool_idx + 1)
        self.assertEqual(win.about_sidebar_btn.maximumHeight(), 38)

    def test_pitc_downloader_timing_defaults(self):
        from pitcsuite.tools.pitc_downloader import PITCPanel
        panel = PITCPanel()
        self.assertEqual(panel.sp_restart.value(), 40)
        self.assertEqual(panel.sp_warmup.value(), 1)
        self.assertEqual(panel.sp_pwait.value(), 5)


    def test_personalization_defaults_and_reset(self):
        panel = SettingsPanel()
        panel.reset_personalization()

        self.assertEqual(panel.theme_cb.currentData(), "vesper")
        self.assertEqual(panel._current_accent, "blue")
        self.assertEqual(panel.font_cb.currentData(), "inter")
        self.assertEqual(panel.scale_cb.currentData(), "standard")
        self.assertEqual(panel.density_cb.currentData(), "compact")
        self.assertEqual(panel.sidebar_style_cb.currentData(), "expanded")
        self.assertEqual(panel.launch_mode_cb.currentData(), "saved")
        self.assertTrue(panel.chk_smooth_anim.isChecked())
        self.assertTrue(panel.chk_auto_update.isChecked())

    def test_installer_script_icon_and_associations(self):
        iss_path = Path(__file__).parent.parent / "installer.iss"
        content = iss_path.read_text(encoding="utf-8")
        self.assertIn("ChangesAssociations=yes", content)
        self.assertIn('Source: "pitcsuite\\assets\\pitcsuite_icon.ico"; DestDir: "{app}"; Flags: ignoreversion', content)
        self.assertIn('IconFilename: "{app}\\pitcsuite_icon.ico"', content)


if __name__ == "__main__":
    unittest.main()

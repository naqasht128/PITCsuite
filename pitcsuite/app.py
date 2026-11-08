import sys

from PySide6.QtWidgets import QApplication

from pitcsuite.icons import app_icon
from pitcsuite.main_window import MainWindow
from pitcsuite.ui_helpers import apply_theme


def _sync_windows_desktop_icon():
    """Ensure desktop shortcut points to the crisp standalone .ico and invalidate Windows Shell icon cache."""
    if sys.platform != "win32":
        return
    try:
        from PySide6.QtCore import QSettings
        settings = QSettings("TNDevelopers", "PITCSuite")
        if settings.value("desktop_icon_synced_v3", False, type=bool):
            return

        import os
        import shutil
        import ctypes
        import subprocess
        from pathlib import Path

        desktop = Path(os.environ.get("USERPROFILE", "")) / "Desktop"
        lnk_path = desktop / "PITC Suite.lnk"
        if not lnk_path.exists():
            settings.setValue("desktop_icon_synced_v3", True)
            return

        assets_ico = Path(__file__).resolve().parent / "assets" / "pitcsuite_icon.ico"
        app_dir = Path(sys.executable).parent if getattr(sys, "frozen", False) else None
        target_ico = None
        if app_dir and app_dir.exists():
            target_ico = app_dir / "pitcsuite_icon.ico"
            if assets_ico.exists() and not target_ico.exists():
                try:
                    shutil.copy2(assets_ico, target_ico)
                except Exception:
                    pass
        if not target_ico or not target_ico.exists():
            if assets_ico.exists():
                target_ico = assets_ico

        if target_ico and target_ico.exists():
            ps_cmd = (
                f"$sh = New-Object -ComObject WScript.Shell; "
                f"$sc = $sh.CreateShortcut('{str(lnk_path)}'); "
                f"if ($sc.IconLocation -ne '{str(target_ico)},0') {{ "
                f"$sc.IconLocation = '{str(target_ico)},0'; $sc.Save(); "
                f"}}"
            )
            subprocess.run(
                ["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", ps_cmd],
                creationflags=0x08000000,
                timeout=3,
                check=False
            )
            ctypes.windll.shell32.SHChangeNotify(0x08000000, 0, None, None)
        settings.setValue("desktop_icon_synced_v3", True)
    except Exception:
        pass


def main():
    _sync_windows_desktop_icon()
    app = QApplication.instance() or QApplication(sys.argv)
    apply_theme(app)
    app.setWindowIcon(app_icon())
    win = MainWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())

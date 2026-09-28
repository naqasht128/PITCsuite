"""Compatibility launcher for the modular PITC Suite package."""

from pitcsuite.app import main
from PySide6.QtWidgets import QApplication
from pitcsuite.license_manager import require_license
import sys
import os
if sys.stdout is None: sys.stdout = open(os.devnull, 'w')
if sys.stderr is None: sys.stderr = open(os.devnull, 'w')


if __name__ == "__main__":
    app = QApplication.instance() or QApplication(sys.argv)
    require_license()
    raise SystemExit(main())


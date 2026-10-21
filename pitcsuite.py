"""Compatibility launcher for the modular PITC Suite package."""

import os
import sys

# Suppress PyMuPDF layout recommendation and prevent missing ONNX model crashes
os.environ.setdefault("PYMUPDF_SUGGEST_LAYOUT_ANALYZER", "0")
sys.modules.setdefault("pymupdf.layout", None)

if sys.stdout is None: sys.stdout = open(os.devnull, 'w')
if sys.stderr is None: sys.stderr = open(os.devnull, 'w')

from pitcsuite.app import main
from PySide6.QtWidgets import QApplication
from pitcsuite.license_manager import require_license


if __name__ == "__main__":
    app = QApplication.instance() or QApplication(sys.argv)
    require_license()
    raise SystemExit(main())

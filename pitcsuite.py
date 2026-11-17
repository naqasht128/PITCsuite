"""Compatibility launcher for the modular PITC Suite package."""

import os
import sys
import warnings

# Suppress PyMuPDF layout recommendation and fitz deprecation warning
os.environ.setdefault("PYMUPDF_SUGGEST_LAYOUT_ANALYZER", "0")
sys.modules.setdefault("pymupdf.layout", None)
warnings.filterwarnings("ignore", message=r".*The `fitz` API is deprecated.*")

# Reconfigure stdout/stderr to utf-8 on Windows to prevent charmap encoding errors
if sys.platform == "win32":
    try:
        if sys.stdout and hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        if sys.stderr and hasattr(sys.stderr, "reconfigure"):
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# Lightweight cv2 stub to prevent ddddocr 1.6.1 from importing 118MB OpenCV
import types
sys.modules.setdefault("cv2", types.ModuleType("cv2"))

if sys.stdout is None: sys.stdout = open(os.devnull, 'w')
if sys.stderr is None: sys.stderr = open(os.devnull, 'w')

from pitcsuite.app import main
from PySide6.QtWidgets import QApplication
from pitcsuite.license_manager import require_license


if __name__ == "__main__":
    app = QApplication.instance() or QApplication(sys.argv)
    require_license()
    raise SystemExit(main())

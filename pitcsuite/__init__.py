import os
import sys

# Suppress PyMuPDF layout recommendation and prevent missing ONNX model crashes
os.environ.setdefault("PYMUPDF_SUGGEST_LAYOUT_ANALYZER", "0")
sys.modules.setdefault("pymupdf.layout", None)

__version__ = "2.8.99"

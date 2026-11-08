import os
from pathlib import Path
from PySide6.QtGui import QIcon, QPainter, QPixmap
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtCore import QByteArray, Qt

ASSETS_DIR = Path(__file__).resolve().parent / "assets"


def app_icon() -> QIcon:
    """Return the application icon with dark roundish background."""
    ico_path = ASSETS_DIR / "pitcsuite_icon.ico"
    if ico_path.exists():
        return QIcon(str(ico_path))
    return QIcon(str(ASSETS_DIR / "pitcsuite_icon.svg"))


def pdf_merger_icon() -> QIcon:
    return get_icon("pdf_merger")


def get_icon(name: str, is_light: bool = None, color: str = None) -> QIcon:
    """
    Load an SVG icon dynamically with crisp contrast:
    - In light mode, transforms dim #9ba8ba lines into crisp #1e293b slate.
    - Supports explicit color tinting (e.g. #ffffff on Emerald Green buttons).
    - Renders high-DPI pixmaps across multiple resolutions.
    """
    svg_path = ASSETS_DIR / f"{name}.svg"
    if not svg_path.exists():
        return QIcon()

    if is_light is None:
        try:
            from pitcsuite.ui_helpers import get_current_theme_key
            is_light = (get_current_theme_key() == "light")
        except Exception:
            is_light = False

    try:
        with open(svg_path, "r", encoding="utf-8") as f:
            svg_data = f.read()

        if color:
            svg_data = svg_data.replace("#9ba8ba", color).replace("#0078d4", color)
        elif is_light:
            svg_data = svg_data.replace("#9ba8ba", "#1e293b")

        renderer = QSvgRenderer(QByteArray(svg_data.encode("utf-8")))
        icon = QIcon()
        for sz in (20, 24, 32, 40, 48, 64):
            pm = QPixmap(sz, sz)
            pm.fill(Qt.transparent)
            p = QPainter(pm)
            p.setRenderHint(QPainter.Antialiasing)
            p.setRenderHint(QPainter.SmoothPixmapTransform)
            renderer.render(p)
            p.end()
            icon.addPixmap(pm)
        return icon
    except Exception:
        return QIcon(str(svg_path))

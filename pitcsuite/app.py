import sys

from PySide6.QtWidgets import QApplication

from pitcsuite.icons import app_icon
from pitcsuite.main_window import MainWindow
from pitcsuite.ui_helpers import apply_theme


def main():
    app = QApplication.instance() or QApplication(sys.argv)
    apply_theme(app)
    app.setWindowIcon(app_icon())
    win = MainWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())

from pathlib import Path
from PySide6.QtGui import QIcon

def app_icon():
    return QIcon(str(Path(__file__).with_name('assets') / 'pitcsuite_icon.svg'))

def pdf_merger_icon():
    return QIcon(str(Path(__file__).with_name('assets') / 'pdf_merger.svg'))

def get_icon(name):
    return QIcon(str(Path(__file__).with_name('assets') / f'{name}.svg'))

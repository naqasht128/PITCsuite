import os, glob
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

hidden_imports = ['openpyxl', 'pypdf', 'fitz', 'pdfplumber', 'pyotp', 'PIL', 'pandas', 'pyautogui', 'requests', 'urllib3', 'reportlab']
hidden_imports += collect_submodules('selenium')
hidden_imports += collect_submodules('webdriver_manager')
hidden_imports += collect_submodules('reportlab')

# Automatically bundle compiled C++ .pyd binaries from Nuitka if present
compiled_binaries = []
for pyd_path in glob.glob('pitcsuite/**/*.pyd', recursive=True):
    rel_folder = os.path.dirname(pyd_path)
    compiled_binaries.append((pyd_path, rel_folder))

a = Analysis(
    ['pitcsuite.py'],
    pathex=[],
    binaries=compiled_binaries,
    datas=[('pitcsuite\\assets', 'pitcsuite\\assets')],
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['pymupdf.layout', 'pymupdf_layout'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='pitcsuite',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['pitcsuite\\assets\\pitcsuite_icon.ico'],
)

# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

hidden_imports = ['openpyxl', 'pypdf', 'fitz', 'pyotp', 'PIL', 'pandas', 'pyautogui', 'requests', 'urllib3']
hidden_imports += collect_submodules('selenium')
hidden_imports += collect_submodules('webdriver_manager')

a = Analysis(
    ['Pitcsuite.py'],
    pathex=[],
    binaries=[],
    datas=[('pitcsuite\\assets', 'pitcsuite\\assets')],
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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

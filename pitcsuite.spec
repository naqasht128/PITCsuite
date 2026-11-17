import os, glob, importlib.util
from PyInstaller.utils.hooks import collect_submodules

hidden_imports = ['openpyxl', 'pypdf', 'fitz', 'pdfplumber', 'pyotp', 'PIL', 'pandas', 'pyautogui', 'requests', 'urllib3', 'reportlab', 'ddddocr', 'onnxruntime']
hidden_imports += collect_submodules('selenium')
hidden_imports += collect_submodules('webdriver_manager')
hidden_imports += collect_submodules('reportlab')
hidden_imports += collect_submodules('ddddocr')

# Bundle both common_old.onnx (default 13.6MB OCR model) and common.onnx (54MB beta model), plus charsets
ddddocr_datas = []
try:
    dd_spec = importlib.util.find_spec('ddddocr')
    if dd_spec and dd_spec.submodule_search_locations:
        dd_path = list(dd_spec.submodule_search_locations)[0]
        for fname in ['common_old.onnx', 'common.onnx', 'charsets.py', 'logo.png']:
            fpath = os.path.join(dd_path, fname)
            if os.path.exists(fpath):
                ddddocr_datas.append((fpath, 'ddddocr'))
except Exception:
    pass

# Automatically bundle compiled C++ .pyd binaries from Nuitka if present
compiled_binaries = []
for pyd_path in glob.glob('pitcsuite/**/*.pyd', recursive=True):
    rel_folder = os.path.dirname(pyd_path)
    compiled_binaries.append((pyd_path, rel_folder))

a = Analysis(
    ['pitcsuite.py'],
    pathex=[],
    binaries=compiled_binaries,
    datas=[('pitcsuite\\assets', 'pitcsuite\\assets')] + ddddocr_datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=['rthook_cv2.py'],
    excludes=['pymupdf.layout', 'pymupdf_layout', 'cv2', 'opencv'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
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

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='pitcsuite',
)

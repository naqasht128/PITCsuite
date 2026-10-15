#!/usr/bin/env python3
"""
PITC Suite — Nuitka Binary Compiler Script
Compiles all proprietary Python modules in pitcsuite/ into native C++ (.pyd) binaries.
Protects source code against decompilers, byte-code extractors, and AI models.
"""

import os
import sys
import glob
import subprocess
from pathlib import Path

# Ensure UTF-8 stdout/stderr on Windows to prevent UnicodeEncodeError
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

ROOT_DIR = Path(__file__).resolve().parent.parent
PITCSUITE_DIR = ROOT_DIR / "pitcsuite"
TOOLS_DIR = PITCSUITE_DIR / "tools"


def get_target_files():
    """Collect all proprietary Python source files to compile into native C++."""
    targets = []

    # Core proprietary modules in pitcsuite/
    core_modules = [
        "license_manager.py",
        "auto_updater.py",
        "main_window.py",
        "ui_helpers.py",
    ]
    for name in core_modules:
        p = PITCSUITE_DIR / name
        if p.exists():
            targets.append(p)

    # All tool modules in pitcsuite/tools/
    if TOOLS_DIR.exists():
        for p in TOOLS_DIR.glob("*.py"):
            if p.name != "__init__.py":
                targets.append(p)

    return sorted(targets)


def compile_with_nuitka():
    if os.environ.get("SKIP_NUITKA") == "1":
        print("[INFO] SKIP_NUITKA=1 detected. Skipping Nuitka compilation and using standard source files.")
        return

    targets = get_target_files()
    print("=" * 60)
    print(f"[SECURE] NUITKA COMPILER: Securing {len(targets)} Python modules to native C++ (.pyd)")
    print("=" * 60)

    success_count = 0
    for idx, target in enumerate(targets, 1):
        rel_path = target.relative_to(ROOT_DIR)
        print(f"\n[{idx}/{len(targets)}] Compiling {rel_path} to native C++...")

        cmd = [
            sys.executable, "-m", "nuitka",
            "--module",
            "--assume-yes-for-downloads",
            f"--output-dir={str(target.parent)}",
            "--remove-output",
            "--no-progressbar",
            str(target)
        ]

        result = subprocess.run(cmd, cwd=str(ROOT_DIR))
        if result.returncode != 0:
            print(f"[ERROR] Failed to compile {rel_path} with Nuitka (exit code {result.returncode})")
            sys.exit(result.returncode)

        # Find the generated .pyd file
        stem = target.stem
        parent_dir = target.parent
        pyd_matches = list(parent_dir.glob(f"{stem}*.pyd"))

        if not pyd_matches:
            print(f"[ERROR] Compiled .pyd for {stem} not found in {parent_dir}")
            sys.exit(1)

        print(f"  [OK] Generated native C++ binary: {pyd_matches[0].name}")

        # Remove the original .py file so only .pyd is packaged by PyInstaller
        target.unlink()
        print(f"  [OK] Removed original source file: {target.name}")
        success_count += 1

    print("\n" + "=" * 60)
    print(f"[DONE] NUITKA COMPILATION COMPLETED: {success_count}/{len(targets)} modules secured!")
    print("=" * 60 + "\n")


if __name__ == "__main__":
    compile_with_nuitka()

#!/usr/bin/env python3
"""
PITC Suite — Automated Release & Update Publisher.
Implements the TN Developers Universal Licensing Blueprint.

One-command release pipeline to:
1. Verify Python syntax across pitcsuite codebase (py_compile)
2. Bump version in pitcsuite/__init__.py and installer.iss
3. Stage and commit pending changes
4. Push commit to GitHub (main branch)
5. Create and push release tag (vX.Y.Z)
6. Automatically trigger GitHub Actions to build PITCSuite_Setup.exe and publish GitHub Release
7. Notify all active installed desktop app users!

Usage:
    python release.py 2.8.92
    python release.py 2.8.92 -m "Added CP 139 improvements and bug fixes"
    python release.py  (Interactive mode)
"""

from __future__ import annotations

import argparse
import os
import py_compile
import re
import subprocess
import sys
from pathlib import Path

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


ROOT_DIR = Path(__file__).resolve().parent
INIT_PY = ROOT_DIR / "pitcsuite" / "__init__.py"
INSTALLER_ISS = ROOT_DIR / "installer.iss"
README_MD = ROOT_DIR / "README.md"
GITHUB_REPO = "naqasht128/PitcSuite"


def run_cmd(cmd: list[str], check: bool = True) -> subprocess.CompletedProcess:
    """Run a system command and return output."""
    print(f"  → Running: {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=str(ROOT_DIR), text=True, capture_output=True)
    if check and result.returncode != 0:
        print(f"\n❌ Error executing command: {' '.join(cmd)}")
        if result.stderr:
            print(f"Error details:\n{result.stderr.strip()}")
        if result.stdout:
            print(f"Output:\n{result.stdout.strip()}")
        sys.exit(result.returncode)
    return result


def get_current_version() -> str:
    """Read __version__ from pitcsuite/__init__.py."""
    if not INIT_PY.exists():
        print(f"❌ Error: {INIT_PY} not found.")
        sys.exit(1)

    content = INIT_PY.read_text(encoding="utf-8")
    match = re.search(r'__version__\s*=\s*["\']([^"\']+)["\']', content)
    if not match:
        print(f"❌ Error: Could not find __version__ in {INIT_PY}")
        sys.exit(1)
    return match.group(1).strip()


def suggest_next_version(current: str) -> str:
    """Suggest next patch version (e.g. 2.8.91 -> 2.8.92)."""
    parts = current.split(".")
    try:
        parts[-1] = str(int(parts[-1]) + 1)
        return ".".join(parts)
    except Exception:
        return f"{current}.1"


def update_version_files(new_version: str):
    """Write new version into pitcsuite/__init__.py, installer.iss, and README.md."""
    # 1. Update pitcsuite/__init__.py
    init_content = INIT_PY.read_text(encoding="utf-8")
    new_init_content, count1 = re.subn(
        r'__version__\s*=\s*["\'][^"\']+["\']',
        f'__version__ = "{new_version}"',
        init_content
    )
    if count1 == 0:
        print(f"❌ Error: Failed to replace __version__ in {INIT_PY}")
        sys.exit(1)
    INIT_PY.write_text(new_init_content, encoding="utf-8")
    print(f"  ✔ Updated pitcsuite/__init__.py: __version__ = \"{new_version}\"")

    # 2. Update installer.iss
    if INSTALLER_ISS.exists():
        iss_content = INSTALLER_ISS.read_text(encoding="utf-8")
        new_iss_content = re.sub(
            r'#define\s+AppVersion\s+["\'][^"\']+["\']',
            f'#define AppVersion "{new_version}"',
            iss_content
        )
        new_iss_content = re.sub(
            r'AppVersion=[0-9A-Za-z\.\-]+',
            f'AppVersion={new_version}',
            new_iss_content
        )
        INSTALLER_ISS.write_text(new_iss_content, encoding="utf-8")
        print(f"  ✔ Updated installer.iss: AppVersion \"{new_version}\"")

    # 3. Update README.md badge if present
    if README_MD.exists():
        readme_content = README_MD.read_text(encoding="utf-8")
        new_readme_content = re.sub(
            r'version-[0-9\.]+-orange',
            f'version-{new_version}-orange',
            readme_content
        )
        if new_readme_content != readme_content:
            README_MD.write_text(new_readme_content, encoding="utf-8")
            print(f"  ✔ Updated README.md version badge: {new_version}")


def verify_python_syntax():
    """Compiles all Python files in the repository to guarantee no syntax errors."""
    print("  Checking syntax of Python files...")
    py_files = [
        ROOT_DIR / "pitcsuite.py",
        ROOT_DIR / "generate_key.py",
        ROOT_DIR / "release.py",
        ROOT_DIR / "scripts" / "compile_nuitka.py"
    ]
    for p in (ROOT_DIR / "pitcsuite").rglob("*.py"):
        py_files.append(p)

    errors = 0
    for file in py_files:
        if file.exists():
            try:
                py_compile.compile(str(file), doraise=True)
            except py_compile.PyCompileError as e:
                print(f"  ❌ Syntax error in {file.name}:\n{e}")
                errors += 1

    if errors > 0:
        print(f"\n❌ Aborting release due to {errors} syntax error(s).")
        sys.exit(1)
    print("  ✔ All Python files passed syntax validation.")


def main():
    parser = argparse.ArgumentParser(description="PITC Suite - Automated Release & Update Publisher")
    parser.add_argument("version", nargs="?", help="New version number (e.g. 2.8.92)")
    parser.add_argument("-m", "--message", default="", help="Release note / commit message")
    parser.add_argument("--skip-tests", action="store_true", help="Skip Python syntax pre-compilation check")
    args = parser.parse_args()

    print("\n" + "=" * 60)
    print("🚀  PITC SUITE — AUTOMATED RELEASE PUBLISHER")
    print("=" * 60)

    current_ver = get_current_version()
    print(f"Current version in codebase: v{current_ver}")

    # Determine target version
    target_ver = args.version
    if not target_ver:
        suggested = suggest_next_version(current_ver)
        user_input = input(f"Enter new version number [default: {suggested}]: ").strip()
        target_ver = user_input if user_input else suggested

    # Clean version string (e.g. strip leading 'v')
    target_ver = target_ver.lstrip("v").strip()
    tag_name = f"v{target_ver}"

    if not re.match(r"^\d+\.\d+(\.\d+)?", target_ver):
        print(f"❌ Error: Invalid version format '{target_ver}'. Must be like 2.8.92 or 3.0.0")
        sys.exit(1)

    print(f"Target release version:    {tag_name}")

    # Release note
    msg = args.message.strip()
    if not msg:
        user_msg = input("Enter release notes / changes summary (optional): ").strip()
        msg = user_msg if user_msg else f"Release {tag_name}"

    # 1. Syntax check
    if not args.skip_tests:
        print("\n[Step 1/5] Verifying Python code syntax...")
        verify_python_syntax()

    # 2. Update version files
    print(f"\n[Step 2/5] Updating version numbers to {target_ver}...")
    update_version_files(target_ver)

    # 3. Git Stage & Commit
    print("\n[Step 3/5] Staging code changes & committing...")
    run_cmd(["git", "add", "."])
    commit_msg = f"Release {tag_name}: {msg}"
    run_cmd(["git", "commit", "-m", commit_msg], check=False)
    print("  ✔ Git commit created.")

    # 4. Push to main branch
    print("\n[Step 4/5] Pushing changes to GitHub main branch...")
    run_cmd(["git", "push", "origin", "main"])
    print("  ✔ Pushed to origin/main.")

    # 5. Create & Push Git Tag
    print(f"\n[Step 5/5] Creating and pushing tag '{tag_name}'...")
    run_cmd(["git", "tag", "-d", tag_name], check=False)
    run_cmd(["git", "tag", "-a", tag_name, "-m", commit_msg])
    run_cmd(["git", "push", "origin", tag_name, "--force"])
    print(f"  ✔ Tag '{tag_name}' pushed to GitHub.")

    # 6. Optional: Register release on Vercel backend
    api_url = os.environ.get("LICENSE_API_URL", "https://pit-csuite.vercel.app").rstrip("/")
    admin_secret = os.environ.get("ADMIN_SECRET", "")
    admin_cfg = ROOT_DIR / ".admin_config.json"
    if admin_cfg.exists():
        try:
            import json
            cfg = json.loads(admin_cfg.read_text(encoding="utf-8"))
            if not admin_secret:
                admin_secret = cfg.get("admin_secret", "")
            if api_url == "https://pit-csuite.vercel.app" and cfg.get("api_url"):
                api_url = cfg.get("api_url").rstrip("/")
        except Exception:
            pass

    if admin_secret:
        try:
            import json
            import urllib.request
            set_url = f"{api_url}/api/admin/set-update"
            payload = json.dumps({
                "version": target_ver,
                "release_notes": msg,
                "file_name": "PITCSuite_Setup.exe"
            }).encode("utf-8")
            req = urllib.request.Request(
                set_url,
                data=payload,
                headers={
                    "Content-Type": "application/json",
                    "x-admin-secret": admin_secret,
                    "User-Agent": "PITCSuite-Release/1.0"
                },
                method="POST"
            )
            with urllib.request.urlopen(req, timeout=5) as resp:
                if resp.status == 200:
                    print(f"  ✔ Update registered with Vercel API: v{target_ver}")
        except Exception as e:
            print(f"  ℹ Notice: Could not register update directly with Vercel API ({e})")

    print("\n" + "=" * 60)
    print("🎉  RELEASE PIPELINE TRIGGERED SUCCESSFULLY!")
    print("=" * 60)
    print(f"\nWhat happens next:")
    print(f"1. GitHub Actions has started building PITCSuite_Setup.exe (Inno Setup Installer).")
    print(f"2. Upon completion, GitHub Actions will publish Release {tag_name} with:")
    print(f"   • PITCSuite_Setup.exe  (Professional Inno Setup Windows Installer)")
    print(f"   • pitcsuite.exe        (Standalone Executable)")
    print(f"3. All desktop clients will automatically receive the update prompt on next launch.")
    print(f"4. Users click 'Install Update' to upgrade seamlessly in-place.\n")
    print("💡 Tip: Track live build progress on GitHub Actions:")
    print(f"   https://github.com/{GITHUB_REPO}/actions\n")


if __name__ == "__main__":
    main()

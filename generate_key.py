#!/usr/bin/env python3
"""
PITC Suite — Admin Serial Key Generator Tool.
Implements the TN Developers Universal Licensing Blueprint.

Generates new subscription serial keys via the Vercel Express API backed by Supabase
and automatically copies them to the clipboard to send directly to buyers.

Supports both a modern VS Code Dark+ Tkinter GUI and CLI command-line mode.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

import urllib.error
import urllib.request

# Default API Configuration
DEFAULT_API_URL = os.environ.get("LICENSE_API_URL", "https://pit-csuite.vercel.app").rstrip("/")
DEFAULT_ADMIN_SECRET = os.environ.get("ADMIN_SECRET", "")
def get_base_dir() -> Path:
    """Returns directory of running script or executable."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

BASE_DIR = get_base_dir()
CONFIG_FILE = BASE_DIR / ".admin_config.json"
HISTORY_FILE = BASE_DIR / "serialkeys_history.txt"

# VS Code Dark+ Palette
BG_DARK = "#1e1e1e"
CARD_BG = "#252526"
BORDER_COLOR = "#383838"
TEXT_WHITE = "#ffffff"
TEXT_MAIN = "#cccccc"
TEXT_MUTED = "#8b949e"
ACCENT_GREEN = "#238636"
ACCENT_GREEN_HOVER = "#2ea043"
INPUT_BG = "#2d2d2d"
BTN_SEC_BG = "#3a3d41"
BTN_SEC_HOVER = "#45494e"
UI_FONT = "Segoe UI" if sys.platform == "win32" else "DejaVu Sans"
MONO_FONT = "Consolas" if sys.platform == "win32" else "DejaVu Sans Mono"

PLANS = [
    ("1_month", "1 Month  —  Rs. 600  (30 Days, 1 Device)"),
    ("3_months", "3 Months  —  Rs. 1,400  (90 Days, Up to 2 Devices)"),
    ("6_months", "6 Months  —  Rs. 2,400  (180 Days, Up to 2 Devices)"),
    ("1_year", "1 Year  —  Rs. 4,000  (365 Days, Up to 3 Devices)"),
]


def load_cached_admin_config() -> tuple[str, str]:
    """Load cached API URL and Secret from local JSON file."""
    if CONFIG_FILE.exists():
        try:
            data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            return data.get("api_url", DEFAULT_API_URL), data.get("admin_secret", DEFAULT_ADMIN_SECRET)
        except Exception:
            pass
    return DEFAULT_API_URL, DEFAULT_ADMIN_SECRET


def save_cached_admin_config(api_url: str, admin_secret: str):
    """Save API URL and Secret to local JSON file for convenience."""
    try:
        data = {"api_url": api_url.strip(), "admin_secret": admin_secret.strip()}
        CONFIG_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except Exception:
        pass


def record_generated_key(serial_key: str, plan_type: str, duration_days: int, max_devices: int):
    """Appends generated key to local audit log."""
    try:
        from datetime import datetime
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{now_str}] KEY: {serial_key} | PLAN: {plan_type} | DAYS: {duration_days} | MAX_DEVICES: {max_devices}\n"
        with open(HISTORY_FILE, "a", encoding="utf-8") as f:
            f.write(line)
    except Exception:
        pass


def generate_key_api(api_url: str, admin_secret: str, plan_type: str) -> dict:
    """Call the remote API to generate a new key using standard library."""
    url = f"{api_url.rstrip('/')}/api/admin/generate"
    headers = {
        "Content-Type": "application/json",
        "x-admin-secret": admin_secret.strip(),
        "User-Agent": "PITCSuite-AdminConsole/1.0"
    }
    payload = json.dumps({"plan_type": plan_type}).encode("utf-8")
    req = urllib.request.Request(url, data=payload, headers=headers, method="POST")

    try:
        with urllib.request.urlopen(req, timeout=12) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        try:
            err_body = json.loads(e.read().decode("utf-8"))
            err_msg = err_body.get("error") or err_body.get("message")
        except Exception:
            err_msg = str(e.reason)
        raise RuntimeError(err_msg or f"HTTP {e.code}")
    except urllib.error.URLError as e:
        raise RuntimeError(f"Connection failed: {e.reason}")
    except Exception as e:
        raise RuntimeError(f"Request failed: {str(e)}")


class AdminConsoleGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PITC Suite — Admin Serial Key Generator")
        self.configure(bg=BG_DARK)
        self.geometry("560x660")
        self.minsize(520, 620)

        self.cached_api, self.cached_secret = load_cached_admin_config()

        self._set_icon()
        self._init_styles()
        self._build_ui()

    def _set_icon(self):
        icon_candidates = [
            BASE_DIR / "pitcsuite" / "assets" / "pitcsuite_icon.ico",
            Path(__file__).resolve().parent / "pitcsuite" / "assets" / "pitcsuite_icon.ico",
            Path(getattr(sys, "_MEIPASS", "")) / "pitcsuite" / "assets" / "pitcsuite_icon.ico",
        ]
        for candidate in icon_candidates:
            if candidate.exists():
                try:
                    self.iconbitmap(str(candidate))
                    break
                except Exception:
                    pass

    def _init_styles(self):
        self.style = ttk.Style(self)
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        self.style.configure("Root.TFrame", background=BG_DARK)
        self.style.configure("Card.TFrame", background=CARD_BG)
        self.style.configure("Title.TLabel", background=BG_DARK, foreground=TEXT_WHITE, font=(UI_FONT, 14, "bold"))
        self.style.configure("Sub.TLabel", background=BG_DARK, foreground=TEXT_MUTED, font=(UI_FONT, 9))
        self.style.configure("CardTitle.TLabel", background=CARD_BG, foreground=TEXT_WHITE, font=(UI_FONT, 10, "bold"))
        self.style.configure("CardText.TLabel", background=CARD_BG, foreground=TEXT_MAIN, font=(UI_FONT, 9))
        self.style.configure("CardMuted.TLabel", background=CARD_BG, foreground=TEXT_MUTED, font=(UI_FONT, 8))
        self.style.configure("TRadiobutton", background=CARD_BG, foreground=TEXT_MAIN, font=(UI_FONT, 9))
        self.style.map("TRadiobutton", background=[("active", CARD_BG)], foreground=[("active", TEXT_WHITE)])

    def _build_ui(self):
        container = ttk.Frame(self, style="Root.TFrame", padding=(24, 20, 24, 20))
        container.pack(fill="both", expand=True)

        # Header
        ttk.Label(container, text="PITC Suite — Serial Key Generator", style="Title.TLabel").pack(anchor="w")
        ttk.Label(container, text="Generate customer license keys backed by Vercel & Supabase", style="Sub.TLabel").pack(anchor="w", pady=(2, 14))

        # API Configuration Card
        config_card = ttk.Frame(container, style="Card.TFrame", padding=(16, 14))
        config_card.pack(fill="x", pady=(0, 14))

        ttk.Label(config_card, text="SERVER CONFIGURATION", style="CardTitle.TLabel").pack(anchor="w", pady=(0, 8))

        # API URL
        ttk.Label(config_card, text="License Server URL:", style="CardText.TLabel").pack(anchor="w")
        self.api_url_var = tk.StringVar(value=self.cached_api)
        api_entry = tk.Entry(
            config_card,
            textvariable=self.api_url_var,
            bg=INPUT_BG,
            fg=TEXT_WHITE,
            insertbackground=TEXT_WHITE,
            relief="flat",
            font=(UI_FONT, 9),
            highlightthickness=1,
            highlightbackground=BORDER_COLOR,
            highlightcolor=ACCENT_GREEN,
        )
        api_entry.pack(fill="x", pady=(4, 10), ipady=4)

        # Admin Secret
        ttk.Label(config_card, text="Admin Secret (x-admin-secret):", style="CardText.TLabel").pack(anchor="w")
        self.secret_var = tk.StringVar(value=self.cached_secret)
        secret_box = ttk.Frame(config_card, style="Card.TFrame")
        secret_box.pack(fill="x", pady=(4, 0))

        self.secret_entry = tk.Entry(
            secret_box,
            textvariable=self.secret_var,
            bg=INPUT_BG,
            fg=TEXT_WHITE,
            show="*",
            insertbackground=TEXT_WHITE,
            relief="flat",
            font=(UI_FONT, 9),
            highlightthickness=1,
            highlightbackground=BORDER_COLOR,
            highlightcolor=ACCENT_GREEN,
        )
        self.secret_entry.pack(side="left", fill="x", expand=True, ipady=4)

        self.show_secret_btn = tk.Button(
            secret_box,
            text="👁 Show",
            bg=BTN_SEC_BG,
            fg=TEXT_MAIN,
            activebackground=BTN_SEC_HOVER,
            activeforeground=TEXT_WHITE,
            relief="flat",
            font=(UI_FONT, 8),
            command=self._toggle_secret,
            padx=8,
            cursor="hand2",
        )
        self.show_secret_btn.pack(side="left", padx=(6, 0), ipady=2)

        # Plan Selection Card
        plan_card = ttk.Frame(container, style="Card.TFrame", padding=(16, 14))
        plan_card.pack(fill="x", pady=(0, 14))

        ttk.Label(plan_card, text="SELECT SUBSCRIPTION PLAN", style="CardTitle.TLabel").pack(anchor="w", pady=(0, 10))

        self.plan_var = tk.StringVar(value="1_month")
        for code, label in PLANS:
            r = ttk.Radiobutton(plan_card, text=label, value=code, variable=self.plan_var, style="TRadiobutton")
            r.pack(anchor="w", pady=3)

        # Generate Button
        self.gen_btn = tk.Button(
            container,
            text="⚡  GENERATE NEW SERIAL KEY",
            bg=ACCENT_GREEN,
            fg=TEXT_WHITE,
            activebackground=ACCENT_GREEN_HOVER,
            activeforeground=TEXT_WHITE,
            relief="flat",
            font=(UI_FONT, 10, "bold"),
            command=self._on_generate,
            cursor="hand2",
            pady=8,
        )
        self.gen_btn.pack(fill="x", pady=(0, 14))

        # Result Display Card
        res_card = ttk.Frame(container, style="Card.TFrame", padding=(16, 14))
        res_card.pack(fill="both", expand=True)

        ttk.Label(res_card, text="GENERATED SERIAL KEY", style="CardTitle.TLabel").pack(anchor="w", pady=(0, 6))

        self.key_var = tk.StringVar(value="")
        key_box = ttk.Frame(res_card, style="Card.TFrame")
        key_box.pack(fill="x", pady=(0, 10))

        self.key_entry = tk.Entry(
            key_box,
            textvariable=self.key_var,
            font=(MONO_FONT, 13, "bold"),
            bg=INPUT_BG,
            fg="#4ade80",
            insertbackground=TEXT_WHITE,
            relief="flat",
            justify="center",
            highlightthickness=1,
            highlightbackground=BORDER_COLOR,
            highlightcolor=ACCENT_GREEN,
        )
        self.key_entry.pack(side="left", fill="x", expand=True, ipady=6)

        self.copy_btn = tk.Button(
            key_box,
            text="📋 Copy",
            bg=BTN_SEC_BG,
            fg=TEXT_WHITE,
            activebackground=BTN_SEC_HOVER,
            activeforeground=TEXT_WHITE,
            relief="flat",
            font=(UI_FONT, 9, "bold"),
            command=self._copy_key,
            padx=12,
            cursor="hand2",
        )
        self.copy_btn.pack(side="left", padx=(8, 0), ipady=4)

        self.status_var = tk.StringVar(value="Ready. Select a plan and click Generate.")
        self.status_lbl = ttk.Label(res_card, textvariable=self.status_var, style="CardMuted.TLabel")
        self.status_lbl.pack(anchor="w")

    def _toggle_secret(self):
        if self.secret_entry.cget("show") == "*":
            self.secret_entry.configure(show="")
            self.show_secret_btn.configure(text="🔒 Hide")
        else:
            self.secret_entry.configure(show="*")
            self.show_secret_btn.configure(text="👁 Show")

    def _copy_key(self):
        k = self.key_var.get().strip()
        if not k:
            return
        self.clipboard_clear()
        self.clipboard_append(k)
        self.copy_btn.configure(text="✓ Copied", bg="#15803d")
        self.after(2000, lambda: self.copy_btn.configure(text="📋 Copy", bg=BTN_SEC_BG))

    def _on_generate(self):
        api_url = self.api_url_var.get().strip()
        secret = self.secret_var.get().strip()
        plan = self.plan_var.get()

        if not api_url:
            messagebox.showwarning("Missing URL", "Please enter the License Server URL.")
            return

        if not secret:
            messagebox.showwarning("Missing Secret", "Please enter the Admin Secret.")
            return

        save_cached_admin_config(api_url, secret)

        self.gen_btn.configure(text="⏳  Generating...", state="disabled")
        self.status_var.set("Connecting to server and creating serial key...")
        self.update_idletasks()

        try:
            res = generate_key_api(api_url, secret, plan)
            key = res.get("serial_key", "")
            duration = res.get("duration_days", 0)
            devices = res.get("max_devices", 1)

            self.key_var.set(key)
            self._copy_key()
            record_generated_key(key, plan, duration, devices)

            self.status_var.set(f"✓ Key generated & copied to clipboard! Plan: {plan} ({duration} days, {devices} devices)")
            messagebox.showinfo(
                "Success",
                f"Serial Key Generated Successfully!\n\n"
                f"Key: {key}\n"
                f"Plan: {res.get('plan_label', plan)}\n"
                f"Devices: Up to {devices}\n\n"
                "✓ Automatically copied to clipboard!"
            )
        except Exception as e:
            self.status_var.set(f"Error: {e}")
            messagebox.showerror("Generation Failed", f"Failed to generate serial key:\n\n{e}")
        finally:
            self.gen_btn.configure(text="⚡  GENERATE NEW SERIAL KEY", state="normal")


def run_cli(args):
    """Run generator in command-line mode."""
    if sys.platform == "win32" and getattr(sys, "frozen", False):
        try:
            import ctypes
            if ctypes.windll.kernel32.AttachConsole(-1):
                sys.stdout = open("CONOUT$", "w", encoding="utf-8")
                sys.stderr = open("CONOUT$", "w", encoding="utf-8")
        except Exception:
            pass

    cached_url, cached_secret = load_cached_admin_config()
    api_url = (args.url or cached_url or DEFAULT_API_URL).rstrip("/")
    admin_secret = args.secret or cached_secret or DEFAULT_ADMIN_SECRET

    if not admin_secret:
        admin_secret = input("Enter Admin Secret (x-admin-secret): ").strip()

    if not admin_secret:
        print("❌ Error: Admin Secret is required.")
        sys.exit(1)

    plan_type = args.plan
    if not plan_type:
        print("\nAvailable Plans:")
        for code, label in PLANS:
            print(f"  • {code:<10} - {label}")
        plan_type = input("\nEnter plan [default: 1_month]: ").strip() or "1_month"

    print(f"\nGenerating key via {api_url}...")
    try:
        res = generate_key_api(api_url, admin_secret, plan_type)
        key = res.get("serial_key")
        duration = res.get("duration_days")
        devices = res.get("max_devices")

        record_generated_key(key, plan_type, duration, devices)

        print("\n" + "=" * 50)
        print("🎉  SERIAL KEY GENERATED SUCCESSFULLY")
        print("=" * 50)
        print(f"  Serial Key : {key}")
        print(f"  Plan       : {res.get('plan_label', plan_type)}")
        print(f"  Duration   : {duration} Days")
        print(f"  Max Devices: {devices}")
        print("=" * 50)

        # Attempt to copy to Windows clipboard via powershell or tkinter
        try:
            import subprocess
            subprocess.run(["clip"], input=key.encode("utf-8"), check=False)
            print("  ✓ Copied to clipboard automatically!\n")
        except Exception:
            pass

    except Exception as e:
        print(f"\n❌ Error: {e}\n")
        sys.exit(1)


def main():
    parser = argparse.ArgumentParser(description="PITC Suite - Admin Serial Key Generator")
    parser.add_argument("--cli", action="store_true", help="Run in command-line interface mode")
    parser.add_argument("--plan", choices=["1_month", "3_months", "6_months", "1_year"], help="Subscription plan")
    parser.add_argument("--secret", help="Admin secret password")
    parser.add_argument("--url", help="License server URL")

    args = parser.parse_args()

    if args.cli or args.plan:
        run_cli(args)
    else:
        app = AdminConsoleGUI()
        app.mainloop()


if __name__ == "__main__":
    main()

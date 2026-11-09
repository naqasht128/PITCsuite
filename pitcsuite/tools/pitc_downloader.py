import os
import re
import time
import base64
import threading
import configparser
import subprocess

from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton,
    QLineEdit, QFileDialog, QCheckBox, QTextEdit, QProgressBar,
    QVBoxLayout, QHBoxLayout, QGridLayout, QGroupBox, QMessageBox, QComboBox,
    QSpinBox
)
from PySide6.QtCore import Signal, QObject, QThread, QSettings, Qt

from pitcsuite.ui_helpers import lbl, hline, START_BTN, STOP_BTN, notify
from pitcsuite.templates import download_template
from pitcsuite.config import config_path as _config_path

REPORT_HREF = {
    "meter":      "MeterReadingAndConsumption",
    "billing":    "BillingDetails",
    "adjustment": "BillingAdjustmentDetails",
    "payment":    "PaymentDetails",
    "summary":    "CustomerDataSummary",
}

# PITC billing values can differ from the Excel amount in the final rupee
# digit.  Accept differences of up to 9 rupees when locating and highlighting
# the requested amount.
AMOUNT_DIFFERENCE_TOLERANCE = 9


def kill_profile_chrome(profile_dir):
    """Terminate lingering chrome/chromedriver processes associated with profile_dir."""
    if not profile_dir:
        return
    norm_profile = os.path.normpath(profile_dir).lower()
    ps_cmd = (
        "$procs = Get-CimInstance Win32_Process | "
        "Where-Object { ($_.Name -like 'chrome*' -or $_.Name -like 'chromedriver*') -and $_.CommandLine }; "
        "foreach ($p in $procs) { Write-Output \"$($p.ProcessId)|$($p.CommandLine)\" }"
    )
    try:
        res = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_cmd],
            capture_output=True,
            text=True,
            timeout=8,
        )
        for line in res.stdout.splitlines():
            if "|" in line:
                pid_str, cmdline = line.split("|", 1)
                if norm_profile in cmdline.lower():
                    pid = pid_str.strip()
                    if pid.isdigit():
                        subprocess.run(["taskkill", "/F", "/PID", pid], capture_output=True, timeout=5)
    except Exception:
        pass


def cleanup_profile_locks(profile_dir):
    """Remove stale Chrome singleton locks from profile_dir to allow clean reopen."""
    if not profile_dir or not os.path.isdir(profile_dir):
        return
    lock_names = {"SingletonLock", "SingletonCookie", "SingletonSocket", "DevToolsActivePort", "lockfile"}
    try:
        for entry in os.scandir(profile_dir):
            if entry.name in lock_names:
                try:
                    os.remove(entry.path)
                except Exception:
                    pass
    except Exception:
        pass


class PITCWorker(QThread):
    progress = Signal(int)
    log      = Signal(str)

    def __init__(self, excel, report_key, download_dir, profile_dir,
                 restart_after, warmup, print_wait, print_x, print_y,
                 page_from=None, page_to=None, match_mode="first", disco="PESCO"):
        super().__init__()
        self.excel         = excel
        self.report_key    = report_key
        self.download_dir  = download_dir
        self.profile_dir   = profile_dir
        self.restart_after = restart_after
        self.warmup        = warmup
        self.print_wait    = print_wait
        self.print_x       = print_x
        self.print_y       = print_y
        self.page_from     = page_from
        self.page_to       = page_to
        self.match_mode    = match_mode
        self.disco         = disco or "PESCO"
        self.running       = True
        self.stopped       = False
        self._driver       = None
        self._driver_lock  = threading.Lock()

    def stop(self):
        self.running = False
        self.stopped = True
        # Closing Chrome from a short-lived helper prevents a Selenium call in
        # the worker from keeping the application UI hostage.  The GUI thread
        # never waits for this operation.
        with self._driver_lock:
            driver = self._driver
        if driver:
            threading.Thread(target=self._close_driver, args=(driver,), daemon=True).start()

    @staticmethod
    def _close_driver(driver):
        if not driver:
            return
        try:
            if hasattr(driver, "service") and driver.service and getattr(driver.service, "process", None):
                pid = driver.service.process.pid
                try:
                    subprocess.run(["taskkill", "/F", "/T", "/PID", str(pid)], capture_output=True, timeout=5)
                except Exception:
                    pass
            driver.quit()
        except Exception:
            pass

    def run(self):
        try:
            self._run_impl()
        finally:
            with self._driver_lock:
                driver = self._driver
                self._driver = None
            if driver:
                self._close_driver(driver)
            kill_profile_chrome(self.profile_dir)
            cleanup_profile_locks(self.profile_dir)

    def _run_impl(self):
        try:
            self._run_impl_inner()
        except Exception as e:
            import traceback as _tb, os as _os
            with open(_os.path.join(_os.environ['USERPROFILE'], 'Desktop', 'pitc_crash.txt'), 'w') as f: f.write(_tb.format_exc())
            self.log.emit(f"CRASH: {e}")

    def _run_impl_inner(self):
        try:
            from selenium import webdriver
            from selenium.webdriver.chrome.service import Service
            from selenium.webdriver.chrome.options import Options
            from selenium.webdriver.common.by import By
            from selenium.webdriver.common.keys import Keys
            from selenium.webdriver.support.ui import WebDriverWait
            from selenium.webdriver.support import expected_conditions as EC
            from selenium.common.exceptions import (
                InvalidSessionIdException,
                StaleElementReferenceException,
                TimeoutException,
                WebDriverException,
                SessionNotCreatedException,
            )
            from webdriver_manager.chrome import ChromeDriverManager
            from openpyxl import load_workbook
            from pypdf import PdfReader, PdfWriter
            import sys as _sys
            _sys.modules.setdefault("pymupdf.layout", None)
            import pyautogui, fitz, shutil, traceback
            import datetime as _dt
            import pyotp
            import urllib.request
            import email.utils
        except Exception as e:
            import traceback as _tb, os as _os
            with open(_os.path.join(_os.environ['USERPROFILE'], 'Desktop', 'pitc_crash.txt'), 'w') as f: f.write(_tb.format_exc())
            self.log.emit(f"ERROR: Missing library — {e}")
            return

        cfg = configparser.ConfigParser()
        cfg.read(_config_path())
        try:
            USERNAME   = cfg["credentials"]["username"]
            PASSWORD   = cfg["credentials"]["password"]
            OTP_SECRET = re.sub(r"[^A-Za-z2-7]", "", cfg["credentials"].get("otp_secret", "")).upper()
        except Exception:
            self.log.emit("ERROR: config.ini missing or invalid. Set credentials in ⚙️ Settings first.")
            return

        dl = self.download_dir
        os.makedirs(dl, exist_ok=True)
        os.makedirs(self.profile_dir, exist_ok=True)
        pyautogui.FAILSAFE = False
        pyautogui.PAUSE = 0.3

        # ── OTP ──
        server_offset = [0.0]
        offset_calibrated = [False]

        def get_server_time_offset():
            for target_url in ("https://cp.pitc.com.pk", "https://www.google.com"):
                try:
                    t0 = time.time()
                    req = urllib.request.Request(
                        target_url,
                        headers={"User-Agent": "Mozilla/5.0"},
                        method="HEAD"
                    )
                    with urllib.request.urlopen(req, timeout=3) as resp:
                        t1 = time.time()
                        date_str = resp.headers.get("Date")
                        if date_str:
                            server_dt = email.utils.parsedate_to_datetime(date_str)
                            server_ts = server_dt.timestamp()
                            local_ts = (t0 + t1) / 2
                            return server_ts - local_ts
                except Exception:
                    continue
            return 0.0

        def calibrate_clock(force=False):
            if force or not offset_calibrated[0]:
                off = get_server_time_offset()
                server_offset[0] = off
                offset_calibrated[0] = True
                if abs(off) > 1.0:
                    self.log.emit(f"Clock offset calibrated: {off:+.1f}s relative to portal")

        def generate_otp(step_offset=0):
            clean_secret = re.sub(r"[^A-Za-z2-7]", "", OTP_SECRET).upper()
            if not clean_secret:
                return ""
            calibrate_clock()
            totp = pyotp.TOTP(clean_secret)
            server_time = time.time() + server_offset[0]
            remaining = totp.interval - (server_time % totp.interval)

            # If current window expires in < 1.5s, wait briefly for fresh window so OTP doesn't expire in transit
            if remaining < 1.5:
                time.sleep(remaining + 0.2)
                server_time = time.time() + server_offset[0]

            target_time = server_time + (step_offset * totp.interval)
            return totp.at(target_time)

        # ── PDF helpers ──
        def wait_for_new_pdf(existing, timeout=30):
            start = time.time()
            while self.running and time.time() - start < timeout:
                pdfs = {f for f in os.listdir(dl)
                        if f.lower().endswith(".pdf") and not f.lower().endswith(".crdownload")}
                diff = pdfs - existing
                if diff:
                    path = os.path.join(dl, diff.pop())
                    time.sleep(0.4)
                    try:
                        size = os.path.getsize(path)
                    except OSError:
                        size = 0
                    if size < 5000:
                        time.sleep(0.5)
                        try:
                            size = os.path.getsize(path)
                        except OSError:
                            size = 0
                    if size < 5000:
                        raise Exception("Invalid PDF (login/session issue)")
                    return path
                if self.running:
                    time.sleep(0.2)
            if not self.running:
                raise InterruptedError("Download stopped by user")
            raise Exception("PDF not downloaded within timeout")

        def get_unique_name(base_name):
            base_name = str(base_name).strip()
            name = f"{base_name}.pdf"
            if not os.path.exists(os.path.join(dl, name)): return name
            i = 1
            while True:
                name = f"{base_name}_{i}.pdf"
                if not os.path.exists(os.path.join(dl, name)): return name
                i += 1

        def format_month(val):
            if isinstance(val, _dt.datetime):
                return val.strftime("%b - %Y")
            s = str(val).strip()
            for fmt in ("%b-%y", "%b-%Y", "%m-%y", "%m-%Y", "%b %y", "%b %Y", "%m/%y", "%m/%Y", "%d-%m-%Y", "%Y-%m"):
                try:
                    return _dt.datetime.strptime(s, fmt).strftime("%b - %Y")
                except:
                    pass
            return s

        def amount_as_integer(value):
            """Convert an amount to whole rupees, ignoring decimal places."""
            if value is None:
                return None
            if isinstance(value, (int, float)):
                try:
                    return int(round(value))
                except Exception:
                    return None
            try:
                from decimal import Decimal, InvalidOperation
                raw = str(value).replace("\u00a0", " ").strip()
                # PDF text may contain spaces inside a formatted number, e.g. "9 060" or "9, 060"
                raw = re.sub(r"(?<=\d)\s+(?=[\d,])", "", raw)
                # Handle accounting negative in parentheses e.g. "(9,060.00)" or trailing minus "9,060-"
                if raw.startswith("(") and raw.endswith(")"):
                    raw = "-" + raw[1:-1]
                elif raw.endswith("-"):
                    raw = "-" + raw[:-1]
                match = re.search(r"[-+]?\d[\d,]*(?:\.\d+)?", raw)
                if not match:
                    return None
                return int(round(float(Decimal(match.group(0).replace(",", "")))))
            except (InvalidOperation, ValueError, TypeError):
                return None

        def amount_matches(value, amount, tolerance=AMOUNT_DIFFERENCE_TOLERANCE, allow_abs=True):
            """Return True when whole-rupee amounts differ by at most tolerance."""
            v = amount_as_integer(value)
            t = amount_as_integer(amount)
            if v is None or t is None:
                return False
            if abs(v - t) <= tolerance:
                return True
            if allow_abs and abs(abs(v) - abs(t)) <= tolerance:
                return True
            return False

        def amount_rects(page, amount):
            target = amount_as_integer(amount)
            if target is None:
                return []
            rects = []
            for word in page.get_text("words"):
                if amount_matches(word[4], target):
                    rects.append(fitz.Rect(word[:4]))
            if rects:
                return rects

            # Fallback: search for exact text patterns via PyMuPDF search engine
            abs_t = abs(target)
            formatted = f"{abs_t:,}"
            plain = str(abs_t)
            candidates = [
                f"{formatted}.00", formatted,
                f"{plain}.00", plain,
                f"-{formatted}.00", f"-{formatted}",
                f"({formatted}.00)", f"({formatted})",
            ]
            for c in candidates:
                found = page.search_for(c)
                if found:
                    return found
            return []

        def text_has_amount(text, amount):
            target = amount_as_integer(amount)
            if target is None:
                return False
            abs_t = abs(target)
            formatted = f"{abs_t:,}"
            plain = str(abs_t)
            candidates = [
                f"{formatted}.00", formatted,
                f"{plain}.00", plain,
                f"-{formatted}.00", f"-{formatted}",
                f"({formatted}.00)", f"({formatted})",
            ]
            if any(c.lower() in text.lower() for c in candidates):
                return True

            clean = re.sub(r"(?<=\d)[ \t]+(?=\d{3}\b)", "", text.replace("\u00a0", " "))
            clean = re.sub(r"(?<=,)[ \t]+(?=\d)", "", clean)
            tokens = re.findall(r"[-+]?\(?\d[\d,]*(?:\.\d+)?\)?", clean)
            return any(amount_matches(token, target) for token in tokens)

        def debit_amount_in_month_row(page, month, amount):
            """Find the closest debit amount on the row containing the month."""
            target = amount_as_integer(amount)
            if target is None:
                return []

            month_rects = page.search_for(format_month(month))
            if not month_rects:
                return []
            month_rect = month_rects[0]
            month_y = (month_rect.y0 + month_rect.y1) / 2

            candidates = []
            for word in page.get_text("words"):
                word_rect = fitz.Rect(word[:4])
                word_y = (word_rect.y0 + word_rect.y1) / 2
                if word_rect.intersects(month_rect):
                    continue
                if abs(word_y - month_y) > max(10, month_rect.height * 2.0):
                    continue
                value = amount_as_integer(word[4])
                if value is None:
                    continue
                diff = min(abs(value - target), abs(abs(value) - abs(target)))
                candidates.append((diff, word_rect))

            if not candidates:
                return []

            candidates.sort(key=lambda item: item[0])
            return [candidates[0][1]] if candidates[0][0] <= AMOUNT_DIFFERENCE_TOLERANCE else []

        def filter_pdf(pdf_path, month=None, amount=None, page_from=None, page_to=None):
            """Filter by an explicit page range or by month/amount."""
            temp_filtered = pdf_path.replace(".pdf", "_filtered.pdf")

            if page_from is not None and page_to is not None:
                page_doc = fitz.open(pdf_path)
                total = len(page_doc)
                if page_from < 1 or page_to < page_from or page_to > total:
                    page_doc.close()
                    raise Exception(
                        f"Page range {page_from}-{page_to} is outside the downloaded PDF (1-{total})"
                    )
                page_doc.select(list(range(page_from - 1, page_to)))
                page_doc.save(temp_filtered, garbage=4, deflate=True, clean=True)
                page_doc.close()
                os.replace(temp_filtered, pdf_path)
                self.log.emit(f"PDF filtered to pages {page_from}-{page_to}")
                return

            reader = PdfReader(pdf_path)
            fitz_doc = fitz.open(pdf_path)
            writer = PdfWriter()
            month_str = format_month(month).lower() if month else None
            has_month = bool(month_str)
            has_amount = amount_as_integer(amount) is not None
            matching_pages = []
            for page_number in range(len(fitz_doc)):
                page = fitz_doc[page_number]
                text = page.get_text("text").lower()
                m = has_month and month_str in text
                a = has_amount and (text_has_amount(text, amount) or bool(amount_rects(page, amount)))
                matches = (m and a) if (has_month and has_amount) else (m or a)
                if matches:
                    matching_pages.append(page_number)
            fitz_doc.close()
            if self.match_mode == "second":
                selected_pages = matching_pages[1:2]
            elif self.match_mode == "all":
                selected_pages = matching_pages
            else:
                selected_pages = matching_pages[:1]

            if not selected_pages:
                if self.match_mode == "second" and matching_pages:
                    raise Exception("Only one matching page found; second match does not exist")
                raise Exception("No matching page found (month/amount)")

            for page_number in selected_pages:
                writer.add_page(reader.pages[page_number])
            with open(temp_filtered, "wb") as f:
                writer.write(f)
            self.log.emit(
                f"Selected {self.match_mode} matching page(s): "
                f"{len(selected_pages)} of {len(matching_pages)} found"
            )
            # highlight
            doc = fitz.open(temp_filtered)
            for page in doc:
                if has_month:
                    month_rects = page.search_for(format_month(month))
                    for rect in month_rects:
                        page.add_highlight_annot(rect).update()
                if has_amount:
                    areas = []
                    if has_month:
                        areas = debit_amount_in_month_row(page, month, amount)
                    if not areas:
                        areas = amount_rects(page, amount)
                    for rect in areas:
                        page.add_highlight_annot(rect).update()
            highlighted = pdf_path.replace(".pdf", "_highlighted.pdf")
            doc.save(highlighted, garbage=4, deflate=True, clean=True)
            doc.close()
            os.remove(pdf_path)
            os.remove(temp_filtered)
            shutil.move(highlighted, pdf_path)
            self.log.emit("PDF filtered and highlighted")

        def stamp_sr(pdf_path, sr):
            doc = fitz.open(pdf_path)
            total = len(doc)
            for i, page in enumerate(doc, 1):
                # page.rect is rotation-aware, so the stamp stays at the
                # same visual corner on portrait and landscape pages.
                rect = page.rect
                label = f"Sr No. {sr} ({i}/{total})"
                label_width = fitz.get_text_length(label, fontname="hebo", fontsize=14)
                page.insert_text(
                    (rect.width - label_width - 20, 24),
                    label,
                    fontsize=14,
                    fontname="hebo",
                    color=(0, 0, 0),
                )
            stamped = pdf_path.replace(".pdf", "_stamped.pdf")
            doc.save(stamped, garbage=4, deflate=True, clean=True)
            doc.close()
            os.replace(stamped, pdf_path)

        try:
            chromedriver_bin = ChromeDriverManager().install()
        except Exception:
            chromedriver_bin = None

        def start_driver():
            cleanup_profile_locks(self.profile_dir)
            options = Options()
            options.add_argument("--start-maximized")
            options.add_argument("--disable-notifications")
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-background-mode")
            options.add_argument("--disable-features=CalculateNativeWinOcclusion")
            options.add_argument("--no-first-run")
            options.add_argument("--no-default-browser-check")
            options.add_argument("--disable-session-crashed-bubble")
            options.add_argument(f"--user-data-dir={self.profile_dir}")
            options.add_experimental_option("prefs", {
                "download.default_directory": dl,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": False,
            })
            svc = Service(chromedriver_bin) if chromedriver_bin else Service(ChromeDriverManager().install())
            drv = webdriver.Chrome(service=svc, options=options)
            with self._driver_lock:
                self._driver = drv
            return drv, WebDriverWait(drv, 40)

        def do_login(drv, wt):
            selected_disco = getattr(self, "disco", "PESCO")
            self.log.emit(f"Opening cp.pitc.com.pk (DISCO: {selected_disco})…")
            drv.get("https://cp.pitc.com.pk")
            time.sleep(1)

            # Check if session was restored from the persistent profile
            try:
                if drv.find_elements(By.NAME, "ctl00$ContentPlaceHolder1$txtReferenceNumber"):
                    self.log.emit("Existing session active ✔")
                    return
            except Exception:
                pass

            # 1. Select the DISCO from dropdown
            disco_selected = False
            try:
                disco_btn = wt.until(EC.element_to_be_clickable(
                    (By.XPATH, "//button[@data-id='ctl00_ContentPlaceHolder1_cmbSelectDisco']")
                ))
                disco_btn.click()
                time.sleep(0.5)
                disco_opt = wt.until(EC.element_to_be_clickable(
                    (By.XPATH, f"//div[contains(@class,'dropdown-menu')]//span[normalize-space(text())='{selected_disco}']")
                ))
                disco_opt.click()
                disco_selected = True
                self.log.emit(f"Selected DISCO: {selected_disco} ✔")
            except Exception as ex:
                self.log.emit(f"Selecting DISCO via dropdown: {ex}, attempting select fallback…")
                try:
                    from selenium.webdriver.support.ui import Select
                    select_elem = drv.find_element(By.ID, "ctl00_ContentPlaceHolder1_cmbSelectDisco")
                    Select(select_elem).select_by_visible_text(selected_disco)
                    drv.execute_script("""
                        var sel = arguments[0];
                        if (window.showDiv) { showDiv(sel); }
                        if (window.$ && $.fn.selectpicker) { $(sel).selectpicker('render'); }
                    """, select_elem)
                    disco_selected = True
                    self.log.emit(f"Selected DISCO via select element: {selected_disco} ✔")
                except Exception as ex2:
                    self.log.emit(f"Warning: could not select DISCO: {ex2}")

            time.sleep(0.5)

            # 2. Fill Username and Password
            user_input = wt.until(EC.presence_of_element_located((By.NAME, "ctl00$ContentPlaceHolder1$txtUsername")))
            user_input.clear()
            user_input.send_keys(USERNAME)

            pass_input = wt.until(EC.presence_of_element_located((By.NAME, "ctl00$ContentPlaceHolder1$txtPassword")))
            pass_input.clear()
            pass_input.send_keys(PASSWORD)

            wt.until(EC.element_to_be_clickable((By.ID, "btnLogin"))).click()
            if OTP_SECRET:
                self.log.emit("Waiting for OTP screen…")
                step_offsets = [0, 0, -1, 1, 0]
                for attempt in range(5):
                    if not self.running:
                        return
                    try:
                        otp_box = WebDriverWait(drv, 15, poll_frequency=0.2).until(
                            EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_txtOtp"))
                        )
                        drv.execute_script("arguments[0].scrollIntoView({behavior: 'instant', block: 'center'});", otp_box)

                        step_off = step_offsets[attempt]
                        otp_code = generate_otp(step_offset=step_off)
                        self.log.emit(f"Entering OTP (attempt {attempt+1}/5): {otp_code}")

                        otp_box.click()
                        time.sleep(0.05)
                        otp_box.send_keys(Keys.CONTROL, "a")
                        otp_box.send_keys(Keys.BACKSPACE)
                        otp_box.send_keys(str(otp_code).strip())

                        if otp_box.get_attribute("value") != str(otp_code).strip():
                            drv.execute_script("""
                                arguments[0].value = arguments[1];
                                arguments[0].dispatchEvent(new Event('input', {bubbles: true}));
                                arguments[0].dispatchEvent(new Event('change', {bubbles: true}));
                            """, otp_box, str(otp_code).strip())

                        verify_btn = WebDriverWait(drv, 10, poll_frequency=0.2).until(
                            EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_btnVerify"))
                        )
                        try:
                            verify_btn.click()
                        except Exception:
                            drv.execute_script("arguments[0].click();", verify_btn)

                        try:
                            WebDriverWait(drv, 8, poll_frequency=0.2).until(
                                EC.presence_of_element_located((By.NAME, "ctl00$ContentPlaceHolder1$txtReferenceNumber"))
                            )
                            if drv.find_elements(By.NAME, "ctl00$ContentPlaceHolder1$txtReferenceNumber"):
                                self.log.emit("OTP verified ✔")
                                return
                        except TimeoutException:
                            pass

                        calibrate_clock(force=True)
                        self.log.emit(f"OTP attempt {attempt+1} not accepted; retrying…")
                        time.sleep(0.5)
                    except TimeoutException:
                        if not self.running:
                            return
                        calibrate_clock(force=True)
                        self.log.emit(f"OTP input field timed out on attempt {attempt+1}; retrying…")
                    except InvalidSessionIdException:
                        raise
                    except Exception as e:
                        if not self.running:
                            return
                        calibrate_clock(force=True)
                        self.log.emit(f"OTP retry: {e}")
                    time.sleep(0.5)
                if self.running:
                    raise Exception("OTP verification failed after 5 attempts")

        def ensure_logged_in(drv, wt):
            try:
                if drv.find_elements(By.NAME, "ctl00$ContentPlaceHolder1$txtReferenceNumber"):
                    return
            except Exception:
                pass
            self.log.emit("Session expired → re-login")
            do_login(drv, wt)

        def login_with_session_recovery(drv, wt):
            try:
                do_login(drv, wt)
                return drv, wt
            except (InvalidSessionIdException, WebDriverException) as ex:
                self.log.emit(f"Browser session lost during login ({ex}); restarting Chrome")
                self._close_driver(drv)
                kill_profile_chrome(self.profile_dir)
                cleanup_profile_locks(self.profile_dir)
                time.sleep(1)
                drv, wt = start_driver()
                do_login(drv, wt)
                return drv, wt

        def restart_browser(current_driver):
            self.log.emit("Closing browser session…")
            if current_driver:
                self._close_driver(current_driver)
            with self._driver_lock:
                self._driver = None

            time.sleep(2)
            kill_profile_chrome(self.profile_dir)
            cleanup_profile_locks(self.profile_dir)
            time.sleep(1)

            last_err = None
            for attempt in range(1, 4):
                try:
                    self.log.emit(f"Reopening Chrome (attempt {attempt}/3)…")
                    new_drv, new_wt = start_driver()
                    new_drv, new_wt = login_with_session_recovery(new_drv, new_wt)
                    new_wt.until(EC.element_to_be_clickable((By.NAME, "ctl00$ContentPlaceHolder1$txtReferenceNumber")))
                    self.log.emit("Browser reopened and session ready ✔")
                    return new_drv, new_wt
                except Exception as e:
                    last_err = e
                    self.log.emit(f"⚠ Restart attempt {attempt} failed ({e}); retrying…")
                    try:
                        self._close_driver(new_drv)
                    except Exception:
                        pass
                    kill_profile_chrome(self.profile_dir)
                    cleanup_profile_locks(self.profile_dir)
                    time.sleep(2 * attempt)

            raise Exception(f"Failed to restart browser after 3 attempts: {last_err}")

        def open_report_link(drv, wt, href_kw):
            """Click the report link with stale-element retry."""
            for attempt in range(5):
                try:
                    link = wt.until(EC.element_to_be_clickable(
                        (By.XPATH, f"//a[contains(@href,'{href_kw}')]")
                    ))
                    drv.execute_script("arguments[0].scrollIntoView({block:'center'});", link)
                    time.sleep(0.5)
                    drv.execute_script("arguments[0].click();", link)
                    return
                except StaleElementReferenceException:
                    self.log.emit(f"Stale element retry {attempt+1}")
                    time.sleep(1)
                except TimeoutException:
                    self.log.emit(f"Timeout retry {attempt+1}")
                    time.sleep(1)
            raise Exception(f"Could not open report link: {href_kw}")

        def missing_ac_status(drv):
            """Return the workbook status when the portal reports no AC data.

            The portal has used several different messages for the same
            condition over time.  Check the visible page text before opening
            a report and again after opening it so a stale/empty report page
            cannot reach the Print button.
            """
            try:
                text = drv.find_element(By.TAG_NAME, "body").text
            except Exception:
                return None

            normalized = re.sub(r"\s+", " ", text).strip().lower()
            missing_phrases = (
                "data not exist",
                "data does not exist",
                "data doesn't exist",
                "no data found",
                "record not found",
                "bill not found",
                "does not belongs",
                "does not belong",
                "wrong ac no",
                "wrong account",
                "invalid account",
            )
            if any(phrase in normalized for phrase in missing_phrases):
                return "data not exist/Wrong AC No."
            return None

        def ac_found_on_page(drv, ac_no):
            """Check that the requested AC number is present in the report."""
            try:
                page = drv.find_element(By.TAG_NAME, "body").text
            except Exception:
                return False

            compact_page = re.sub(r"[^a-z0-9]", "", page.lower())
            compact_ac = re.sub(r"[^a-z0-9]", "", str(ac_no).lower())
            if not compact_ac:
                return False

            candidates = {compact_ac}
            # Excel may expose a numeric AC as 123456.0.  The portal normally
            # renders the underlying integer without the trailing .0.
            if compact_ac.endswith("0") and str(ac_no).strip().endswith(".0"):
                candidates.add(compact_ac[:-1])
            return any(candidate in compact_page for candidate in candidates)

        def wait_for_report_or_missing(drv, href_kw, timeout=8):
            """Wait briefly for either a report link or a missing-AC result."""
            report_xpath = f"//a[contains(@href,'{href_kw}')]"

            def ready(current_driver):
                missing = missing_ac_status(current_driver)
                if missing:
                    return missing
                if current_driver.find_elements(By.XPATH, report_xpath):
                    return "report"
                return False

            try:
                result = WebDriverWait(drv, timeout, poll_frequency=0.25).until(ready)
                return None if result == "report" else result
            except TimeoutException:
                # No report link after the short search window means this AC
                # is unavailable; avoid the old 5 x 40-second retry path.
                return "data not exist/Wrong AC No."

        def return_to_search_page(drv, wt):
            """Immediately close Crystal Report viewer tab/popup and ensure browser is on search page."""
            # 1. Close all extra tabs (Crystal Report viewer) until only 1 tab remains
            try:
                while len(drv.window_handles) > 1:
                    try:
                        drv.switch_to.window(drv.window_handles[-1])
                        drv.close()
                    except Exception:
                        break
                drv.switch_to.window(drv.window_handles[0])
            except Exception:
                pass

            # 2. Check if search box is already present on current page
            try:
                if drv.find_elements(By.NAME, "ctl00$ContentPlaceHolder1$txtReferenceNumber"):
                    return
            except Exception:
                pass

            # 3. If on report details page (BillingDetails.aspx), go back to search page
            try:
                drv.back()
                WebDriverWait(drv, 4, poll_frequency=0.2).until(
                    EC.presence_of_element_located((By.NAME, "ctl00$ContentPlaceHolder1$txtReferenceNumber"))
                )
                return
            except Exception:
                pass

            # 4. Fallback: navigate directly to cp.pitc.com.pk
            try:
                drv.get("https://cp.pitc.com.pk")
                WebDriverWait(drv, 6, poll_frequency=0.2).until(
                    EC.presence_of_element_located((By.NAME, "ctl00$ContentPlaceHolder1$txtReferenceNumber"))
                )
            except Exception:
                pass

        driver, wait = start_driver()
        driver, wait = login_with_session_recovery(driver, wait)

        wb         = load_workbook(self.excel)
        sheet      = wb.active
        total_rows = sheet.max_row - 1
        processed  = 0
        row        = 2
        href_kw    = REPORT_HREF[self.report_key]

        wait.until(EC.element_to_be_clickable((By.NAME, "ctl00$ContentPlaceHolder1$txtReferenceNumber")))
        if self.warmup > 0:
            time.sleep(self.warmup)

        while True:
            if not self.running: break
            ac = sheet[f"A{row}"].value
            if not ac: break
            sr     = sheet[f"B{row}"].value
            month  = sheet[f"C{row}"].value
            amount = sheet[f"D{row}"].value
            status = sheet[f"E{row}"]

            self.log.emit(f"Row {row} | AC: {ac}")

            try:
                ensure_logged_in(driver, wait)
                report_opened = False
                report_tab_closed = False

                box = wait.until(EC.element_to_be_clickable((By.NAME, "ctl00$ContentPlaceHolder1$txtReferenceNumber")))
                box.send_keys(Keys.CONTROL, "a", Keys.BACKSPACE)
                box.send_keys(str(ac))

                wait.until(EC.element_to_be_clickable((By.NAME, "ctl00$ContentPlaceHolder1$btnGo"))).click()
                time.sleep(2)

                missing_status = wait_for_report_or_missing(driver, href_kw)
                if missing_status:
                    status.value = missing_status
                    self.log.emit(f"→ {missing_status}; print skipped")
                else:
                    open_report_link(driver, wait, href_kw)
                    report_opened = True
                    try:
                        WebDriverWait(driver, 15).until(
                            lambda drv: missing_ac_status(drv) or ac_found_on_page(drv, ac)
                        )
                    except TimeoutException:
                        pass

                    missing_status = missing_ac_status(driver)
                    if missing_status or not ac_found_on_page(driver, ac):
                        missing_status = missing_status or "data not exist/Wrong AC No."
                        status.value = missing_status
                        self.log.emit(f"→ {missing_status}; AC not found on report page; print skipped")
                    else:
                        wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_btnPrint"))).click()

                if missing_status:
                    wb.save(self.excel)
                    return_to_search_page(driver, wait)
                    time.sleep(0.5)
                else:
                    existing = {f for f in os.listdir(dl) if f.lower().endswith(".pdf")}
                    time.sleep(self.print_wait)

                    pyautogui.click(self.print_x, self.print_y)
                    time.sleep(0.3)
                    pyautogui.press("down")
                    time.sleep(0.15)
                    for k in ("tab", "tab", "tab", "enter"):
                        pyautogui.press(k)
                        time.sleep(0.08)

                    new_pdf    = wait_for_new_pdf(existing, timeout=30)

                    # Close Crystal Report tab immediately as soon as PDF download completes
                    return_to_search_page(driver, wait)
                    report_tab_closed = True

                    base_name  = sr if sr else ac
                    final_name = get_unique_name(base_name)
                    final_path = os.path.join(dl, final_name)
                    os.rename(new_pdf, final_path)

                    if self.page_from is not None and self.page_to is not None:
                        filter_pdf(final_path, page_from=self.page_from, page_to=self.page_to)
                    elif month or amount_as_integer(amount) is not None:
                        filter_pdf(final_path, month, amount)

                    if sr:
                        stamp_sr(final_path, sr)

                    status.value = "Done"
                    self.log.emit(f"✅ PDF created: {final_name}")

            except Exception as e:
                traceback.print_exc()
                status.value = "Error"
                self.log.emit(f"⚠ ERROR: {e}")
                return_to_search_page(driver, wait)
                time.sleep(1)

            row += 1
            processed += 1
            self.progress.emit(int((processed / max(total_rows, 1)) * 100))

            if processed % self.restart_after == 0:
                self.log.emit(f"Processed {processed} AC Nos — restarting browser session…")
                wb.save(self.excel)
                try:
                    driver, wait = restart_browser(driver)
                    if self.warmup > 0:
                        time.sleep(self.warmup)
                except Exception as ex:
                    self.log.emit(f"❌ Failed to restart browser: {ex}")
                    wb.save(self.excel)
                    raise

        wb.save(self.excel)


class PITCPanel(QWidget):
    REPORT_LABELS = [
        "Meter Reading & Consumption",
        "Billing Detail",
        "Billing Adjustment",
        "Payment Detail",
        "Customer Data Summary",
    ]
    REPORT_KEYS = ["meter", "billing", "adjustment", "payment", "summary"]
    DISCOS = [
        "PESCO", "LESCO", "GEPCO", "FESCO", "IESCO", "MEPCO",
        "HAZECO", "SEPCO", "HESCO", "QESCO", "TESCO"
    ]

    def __init__(self):
        super().__init__()
        self.worker = None
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(4, 4, 4, 4)

        # ── Excel + Template ──
        g1 = QGroupBox("Excel File")
        g1v = QHBoxLayout(g1)
        g1v.setSpacing(8)
        self.excel_ed = QLineEdit()
        self.excel_ed.setPlaceholderText("Select Excel file (*.xlsx)...")
        self.excel_ed.setReadOnly(True)
        b1 = QPushButton("Browse")
        b1.setCursor(Qt.PointingHandCursor)
        b1.clicked.connect(self.pick_excel)
        btn_tpl = QPushButton("Download Template")
        btn_tpl.setCursor(Qt.PointingHandCursor)
        btn_tpl.clicked.connect(self.get_template)
        g1v.addWidget(self.excel_ed, 1)
        g1v.addWidget(b1)
        g1v.addWidget(btn_tpl)
        layout.addWidget(g1)

        # ── Paths and Environment ──
        g2 = QGroupBox("Paths and Environment")
        g2v = QGridLayout(g2)
        g2v.setHorizontalSpacing(10)
        g2v.setVerticalSpacing(8)
        lbl_dl = QLabel("Download Dir:")
        self.dl_ed = QLineEdit(r"D:\PITC_PDFs")
        g2v.addWidget(lbl_dl, 0, 0)
        g2v.addWidget(self.dl_ed, 0, 1)
        btn_browse_dl = QPushButton("Browse")
        btn_browse_dl.clicked.connect(self._browse_download_dir)
        g2v.addWidget(btn_browse_dl, 0, 2)

        lbl_pr = QLabel("Chrome Profile:")
        self.pr_ed = QLineEdit(r"C:\ChromeProfiles\PITC")
        g2v.addWidget(lbl_pr, 1, 0)
        g2v.addWidget(self.pr_ed, 1, 1)
        btn_browse_pr = QPushButton("Browse")
        btn_browse_pr.clicked.connect(self._browse_profile_dir)
        g2v.addWidget(btn_browse_pr, 1, 2)
        g2v.setColumnStretch(1, 1)
        layout.addWidget(g2)

        # ── Report and Filter Settings ──
        g_rep = QGroupBox("Report and Filter Settings")
        g_rep_layout = QGridLayout(g_rep)
        g_rep_layout.setHorizontalSpacing(10)
        g_rep_layout.setVerticalSpacing(8)

        # Row 0: DISCO and Report Type
        g_rep_layout.addWidget(QLabel("Select DISCO:"), 0, 0)
        self.disco_cb = QComboBox()
        self.disco_cb.addItems(self.DISCOS)
        self.disco_cb.setCurrentText("PESCO")
        self.disco_cb.currentTextChanged.connect(self._on_disco_changed)
        g_rep_layout.addWidget(self.disco_cb, 0, 1)

        g_rep_layout.addWidget(QLabel("Report Type:"), 0, 2)
        self.report_cb = QComboBox()
        self.report_cb.addItems(self.REPORT_LABELS)
        default_idx = self.REPORT_LABELS.index("Billing Adjustment") if "Billing Adjustment" in self.REPORT_LABELS else 2
        self.report_cb.setCurrentIndex(default_idx)
        self.report_cb.currentTextChanged.connect(self._on_report_type_changed)
        g_rep_layout.addWidget(self.report_cb, 0, 3)

        # Row 1: Matching Pages and Optional Page Range
        g_rep_layout.addWidget(QLabel("Download:"), 1, 0)
        self.match_cb = QComboBox()
        self.match_cb.addItem("First matching page", "first")
        self.match_cb.addItem("Second matching page", "second")
        self.match_cb.addItem("All matching pages", "all")
        self.match_cb.setToolTip("Controls which pages are kept when month/amount matches multiple pages.")
        g_rep_layout.addWidget(self.match_cb, 1, 1)

        page_range_h = QHBoxLayout()
        page_range_h.setSpacing(6)
        page_range_h.addWidget(QLabel("From:"))
        self.page_from_ed = QLineEdit()
        self.page_from_ed.setPlaceholderText("e.g. 2")
        self.page_from_ed.setFixedWidth(55)
        page_range_h.addWidget(self.page_from_ed)
        page_range_h.addWidget(QLabel("To:"))
        self.page_to_ed = QLineEdit()
        self.page_to_ed.setPlaceholderText("e.g. 4")
        self.page_to_ed.setFixedWidth(55)
        page_range_h.addWidget(self.page_to_ed)
        page_range_h.addStretch()
        g_rep_layout.addWidget(QLabel("Page Range:"), 1, 2)
        g_rep_layout.addLayout(page_range_h, 1, 3)

        g_rep_layout.setColumnStretch(1, 1)
        g_rep_layout.setColumnStretch(3, 1)
        layout.addWidget(g_rep)

        # ── Timing and Print Settings ──
        g3 = QGroupBox("Timing and Print Settings")
        g3h = QGridLayout(g3)
        g3h.setHorizontalSpacing(10)
        g3h.setVerticalSpacing(8)

        def spin(val, mn=1, mx=9999):
            s = QSpinBox()
            s.setRange(mn, mx)
            s.setValue(val)
            s.setFixedWidth(65)
            return s

        g3h.addWidget(QLabel("Restart:"), 0, 0)
        self.sp_restart = spin(40, 1, 200)
        g3h.addWidget(self.sp_restart, 0, 1)

        g3h.addWidget(QLabel("Warmup:"), 0, 2)
        self.sp_warmup = spin(1, 0, 200)
        g3h.addWidget(self.sp_warmup, 0, 3)

        g3h.addWidget(QLabel("Wait (s):"), 0, 4)
        self.sp_pwait = spin(5, 1, 200)
        g3h.addWidget(self.sp_pwait, 0, 5)

        g3h.addWidget(QLabel("Print X:"), 1, 0)
        self.sp_px = spin(30)
        g3h.addWidget(self.sp_px, 1, 1)

        g3h.addWidget(QLabel("Print Y:"), 1, 2)
        self.sp_py = spin(165)
        g3h.addWidget(self.sp_py, 1, 3)

        layout.addWidget(g3)

        # ── Action Buttons ──
        bh = QHBoxLayout()
        self.btn_start = QPushButton("▶  START")
        self.btn_start.setStyleSheet(START_BTN)
        self.btn_stop = QPushButton("■  STOP")
        self.btn_stop.setStyleSheet(STOP_BTN)
        self.btn_stop.setEnabled(False)
        self.btn_start.clicked.connect(self.run)
        self.btn_stop.clicked.connect(self.halt)
        bh.addWidget(self.btn_start)
        bh.addWidget(self.btn_stop)
        bh.addStretch()
        layout.addLayout(bh)

        self.progress = QProgressBar()
        layout.addWidget(self.progress)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setMinimumHeight(65)
        layout.addWidget(self.log_box, 1)

        # Load saved settings
        try:
            settings = QSettings("TNDevelopers", "PITCSuite")
            if not settings.contains("pitc_timing_defaults_v3"):
                self.sp_restart.setValue(40)
                self.sp_warmup.setValue(1)
                self.sp_pwait.setValue(5)
                settings.setValue("pitc_restart", 40)
                settings.setValue("pitc_warmup", 1)
                settings.setValue("pitc_pwait", 5)
                settings.setValue("pitc_timing_defaults_v3", True)
            else:
                if settings.contains("pitc_restart"): self.sp_restart.setValue(int(settings.value("pitc_restart", 40)))
                if settings.contains("pitc_warmup"): self.sp_warmup.setValue(int(settings.value("pitc_warmup", 1)))
                if settings.contains("pitc_pwait"): self.sp_pwait.setValue(int(settings.value("pitc_pwait", 5)))
            if settings.contains("pitc_px"): self.sp_px.setValue(int(settings.value("pitc_px", 30)))
            if settings.contains("pitc_py"): self.sp_py.setValue(int(settings.value("pitc_py", 165)))

            saved_disco = "PESCO"
            if settings.contains("pitc_disco"):
                saved_disco = str(settings.value("pitc_disco")).strip() or "PESCO"
            else:
                cfg = configparser.ConfigParser()
                path = _config_path()
                if os.path.exists(path):
                    cfg.read(path)
                    saved_disco = cfg.get("credentials", "disco", fallback=cfg.get("pitc_downloader", "disco", fallback="PESCO"))
            idx = self.disco_cb.findText(saved_disco)
            if idx >= 0:
                self.disco_cb.setCurrentIndex(idx)

            saved_report = "Billing Adjustment"
            if settings.contains("pitc_report_type"):
                saved_report = str(settings.value("pitc_report_type")).strip() or "Billing Adjustment"
            idx_rep = self.report_cb.findText(saved_report)
            if idx_rep >= 0:
                self.report_cb.setCurrentIndex(idx_rep)
        except Exception:
            pass

    def _on_report_type_changed(self, text):
        report_type = (text or "").strip()
        if not report_type:
            return
        try:
            settings = QSettings("TNDevelopers", "PITCSuite")
            settings.setValue("pitc_report_type", report_type)
        except Exception:
            pass

    def _on_disco_changed(self, text):
        disco = (text or "").strip()
        if not disco:
            return
        try:
            settings = QSettings("TNDevelopers", "PITCSuite")
            settings.setValue("pitc_disco", disco)
        except Exception:
            pass
        try:
            cfg = configparser.ConfigParser()
            path = _config_path()
            if os.path.exists(path):
                cfg.read(path)
            if not cfg.has_section("credentials"):
                cfg.add_section("credentials")
            cfg.set("credentials", "disco", disco)
            with open(path, "w", encoding="utf-8") as f:
                cfg.write(f)
        except Exception:
            pass

    def pick_excel(self):
        f, _ = QFileDialog.getOpenFileName(self, "Excel File", "", "Excel (*.xlsx)")
        if f: self.excel_ed.setText(f)

    def _browse_download_dir(self):
        p = QFileDialog.getExistingDirectory(self, "Select Download Folder", self.dl_ed.text())
        if p: self.dl_ed.setText(p)

    def _browse_profile_dir(self):
        p = QFileDialog.getExistingDirectory(self, "Select Chrome Profile Folder", self.pr_ed.text())
        if p: self.pr_ed.setText(p)

    def get_template(self):
        download_template(
            self,
            headers      = ["AC No", "Sr No", "Month", "Amount", "Status"],
            default_name = "PITC_Downloader_Template.xlsx"
        )

    def run(self):
        # Save settings on run
        try:
            settings = QSettings("TNDevelopers", "PITCSuite")
            settings.setValue("pitc_restart", self.sp_restart.value())
            settings.setValue("pitc_warmup", self.sp_warmup.value())
            settings.setValue("pitc_pwait", self.sp_pwait.value())
            settings.setValue("pitc_px", self.sp_px.value())
            settings.setValue("pitc_py", self.sp_py.value())
            settings.setValue("pitc_disco", self.disco_cb.currentText())
            settings.setValue("pitc_report_type", self.report_cb.currentText())
        except Exception:
            pass

        excel_path = self.excel_ed.text().strip()
        if not excel_path:
            QMessageBox.warning(self, "Excel File Required", "Please select an Excel file before starting.")
            return
        if not os.path.isfile(excel_path):
            QMessageBox.warning(self, "File Not Found", f"The selected Excel file does not exist:\n{excel_path}")
            return

        page_from_text = self.page_from_ed.text().strip()
        page_to_text = self.page_to_ed.text().strip()
        if bool(page_from_text) != bool(page_to_text):
            QMessageBox.warning(self, "Invalid Page Range", "Enter both Pages from and Pages to, or leave both blank.")
            return

        page_from = page_to = None
        if page_from_text:
            try:
                page_from = int(page_from_text)
                page_to = int(page_to_text)
            except ValueError:
                QMessageBox.warning(self, "Invalid Page Range", "Page numbers must be whole numbers.")
                return
            if page_from < 1 or page_to < page_from:
                QMessageBox.warning(self, "Invalid Page Range", "Pages to must be greater than or equal to Pages from.")
                return

        self.btn_start.setEnabled(False); self.btn_stop.setEnabled(True)
        self.log_box.clear()
        self.worker = PITCWorker(
            excel         = self.excel_ed.text(),
            report_key    = self.REPORT_KEYS[self.report_cb.currentIndex()],
            download_dir  = self.dl_ed.text(),
            profile_dir   = self.pr_ed.text(),
            restart_after = self.sp_restart.value(),
            warmup        = self.sp_warmup.value(),
            print_wait    = self.sp_pwait.value(),
            print_x       = self.sp_px.value(),
            print_y       = self.sp_py.value(),
            page_from     = page_from,
            page_to       = page_to,
            match_mode    = self.match_cb.currentData(),
            disco         = self.disco_cb.currentText(),
        )
        self.worker.log.connect(self.log_box.append)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.finished.connect(self.done)
        self.worker.start()

    def halt(self):
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            self.btn_stop.setEnabled(False)
            self.log_box.append("⛔ Stop requested; closing Chrome safely…")

    def done(self):
        self.btn_start.setEnabled(True); self.btn_stop.setEnabled(False)
        if self.worker and self.worker.stopped:
            self.log_box.append("🛑 Download stopped.")
        else:
            notify(self, "Done", "PITC Downloader completed.", folder_path=self.dl_ed.text())


# ──────────────────────────────────────────────────────────────
#  TOOL 2 — E/Bills Downloader (BillSoft.py)
# ──────────────────────────────────────────────────────────────



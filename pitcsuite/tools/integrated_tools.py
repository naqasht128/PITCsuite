import os
import shutil
import threading
from pathlib import Path

from PySide6.QtCore import QThread, Signal
from PySide6.QtWidgets import (
    QFileDialog, QComboBox, QCheckBox, QDoubleSpinBox, QGroupBox, QHBoxLayout, QLabel, QLineEdit,
    QListWidget, QMessageBox, QPushButton, QTextEdit, QProgressBar, QVBoxLayout, QWidget,
)

from pitcsuite.templates import download_template
from pitcsuite.ui_helpers import START_BTN, STOP_BTN, hline, lbl, notify


class CP22TWorker(QThread):
    log = Signal(str)
    done = Signal(str)
    failed = Signal(str)

    def __init__(self, xlsx_path, pdf_paths):
        super().__init__()
        self.xlsx_path = xlsx_path or None
        self.pdf_paths = pdf_paths

    def run(self):
        try:
            from pitcsuite.tools.cp22t_posting_checker import run_check
            output = run_check(self.xlsx_path, self.pdf_paths, self.log.emit)
            self.done.emit(str(output))
        except Exception as exc:
            self.failed.emit(str(exc))


class CP22TCheckerPanel(QWidget):
    def __init__(self):
        super().__init__()
        self.worker = None
        self.pdf_paths = []
        layout = QVBoxLayout(self); layout.setSpacing(10)
        layout.addWidget(lbl("CP22T Posting Checker", bold=True, color="#7eb8f7"))
        layout.addWidget(lbl("Match CP22T PDF transaction records against an optional Excel list.", color="#556"))
        layout.addWidget(hline())

        group = QGroupBox("Inputs")
        form = QVBoxLayout(group)
        row = QHBoxLayout(); row.addWidget(QLabel("Excel (optional):"))
        self.xlsx_ed = QLineEdit(); self.xlsx_ed.setReadOnly(True); self.xlsx_ed.setProperty("preferred_width", 245)
        row.addWidget(self.xlsx_ed, 1)
        browse = QPushButton("Browse"); browse.clicked.connect(self._browse_excel); row.addWidget(browse)
        form.addLayout(row)
        form.addWidget(QLabel("CP22T PDFs:"))
        self.pdf_list = QListWidget(); self.pdf_list.setMinimumHeight(110); form.addWidget(self.pdf_list)
        buttons = QHBoxLayout()
        add = QPushButton("Add PDFs"); add.clicked.connect(self._add_pdfs); buttons.addWidget(add)
        clear = QPushButton("Clear"); clear.clicked.connect(self._clear_pdfs); buttons.addWidget(clear); buttons.addStretch()
        form.addLayout(buttons)
        layout.addWidget(group)

        actions = QHBoxLayout()
        self.start_btn = QPushButton("▶  START"); self.start_btn.setStyleSheet(START_BTN); self.start_btn.clicked.connect(self._start)
        self.stop_btn = QPushButton("■  STOP"); self.stop_btn.setStyleSheet(STOP_BTN); self.stop_btn.setEnabled(False); self.stop_btn.clicked.connect(self._stop)
        actions.addWidget(self.start_btn); actions.addWidget(self.stop_btn); actions.addStretch(); layout.addLayout(actions)
        self.status = QLabel("Status: Idle"); self.status.setStyleSheet("color:#4a90d9;"); layout.addWidget(self.status)
        self.log_box = QTextEdit(); self.log_box.setReadOnly(True); layout.addWidget(self.log_box)

    def _browse_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Excel input", "", "Excel files (*.xlsx *.xlsm);;All files (*.*)")
        if path: self.xlsx_ed.setText(path)

    def _add_pdfs(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "CP22T PDFs", "", "PDF files (*.pdf);;All files (*.*)")
        for path in paths:
            if path not in self.pdf_paths:
                self.pdf_paths.append(path); self.pdf_list.addItem(path)

    def _clear_pdfs(self):
        self.pdf_paths.clear(); self.pdf_list.clear()

    def _start(self):
        if not self.pdf_paths:
            QMessageBox.warning(self, "Missing PDFs", "Please select at least one CP22T PDF."); return
        self.log_box.clear(); self.status.setText("Status: Working...")
        self.start_btn.setEnabled(False); self.stop_btn.setEnabled(True)
        self.worker = CP22TWorker(self.xlsx_ed.text().strip(), list(self.pdf_paths))
        self.worker.log.connect(self._log); self.worker.done.connect(self._done); self.worker.failed.connect(self._failed)
        self.worker.finished.connect(lambda: (self.start_btn.setEnabled(True), self.stop_btn.setEnabled(False)))
        self.worker.start()

    def _stop(self):
        if self.worker and self.worker.isRunning():
            self.worker.requestInterruption(); self.status.setText("Status: Finishing current operation...")

    def _log(self, message): self.log_box.append(message)
    def _done(self, output):
        self.status.setText("Status: Completed"); notify(self, "Completed", f"Output saved to:\n{output}", file_path=output)
    def _failed(self, message):
        self.status.setText("Status: Failed"); notify(self, "Error", message, critical=True)


class BillScraperWorker(QThread):
    progress = Signal(int, int, str, str)
    done = Signal(str, bool)
    failed = Signal(str)

    def __init__(self, path, delay, stop_event):
        super().__init__(); self.path = Path(path); self.delay = delay; self.stop_event = stop_event

    def run(self):
        try:
            from pitcsuite.tools.bill_scraper_logic import PITCBillScraper
            def progress(done, total, row):
                self.progress.emit(done, total, row.get("_account", ""), row.get("scrape_status", ""))
            PITCBillScraper(delay=self.delay).update_workbook(self.path, progress, self.stop_event)
            self.done.emit(str(self.path), self.stop_event.is_set())
        except Exception as exc:
            self.failed.emit(str(exc))


class BillScraperPanel(QWidget):
    def __init__(self):
        super().__init__(); self.worker = None; self.stop_event = threading.Event()
        layout = QVBoxLayout(self); layout.setSpacing(10)
        layout.addWidget(lbl("PITC Bill Scrapper", bold=True, color="#7eb8f7"))
        layout.addWidget(lbl("Update an Excel workbook with live PITC bill details.", color="#556")); layout.addWidget(hline())
        group = QGroupBox("Workbook")
        form = QVBoxLayout(group); row = QHBoxLayout(); row.addWidget(QLabel("Excel file:"))
        self.xlsx_ed = QLineEdit(); self.xlsx_ed.setReadOnly(True); self.xlsx_ed.setProperty("preferred_width", 245); row.addWidget(self.xlsx_ed, 1)
        browse = QPushButton("Browse"); browse.clicked.connect(self._browse); row.addWidget(browse); form.addLayout(row)
        options = QHBoxLayout(); options.addWidget(QLabel("Delay between bills (seconds):"))
        self.delay = QDoubleSpinBox(); self.delay.setRange(0.1, 30); self.delay.setSingleStep(.1); self.delay.setValue(2.0); options.addWidget(self.delay); options.addStretch(); form.addLayout(options)
        layout.addWidget(group)
        actions = QHBoxLayout(); self.start_btn = QPushButton("▶  START"); self.start_btn.setStyleSheet(START_BTN); self.start_btn.clicked.connect(self._start)
        self.stop_btn = QPushButton("■  STOP"); self.stop_btn.setStyleSheet(STOP_BTN); self.stop_btn.setEnabled(False); self.stop_btn.clicked.connect(self._stop)
        actions.addWidget(self.start_btn); actions.addWidget(self.stop_btn); actions.addStretch(); layout.addLayout(actions)
        self.progress = QLabel("Processed: 0"); layout.addWidget(self.progress)
        self.progress_bar = QProgressBar(); self.progress_bar.setRange(0, 1); self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)
        self.log_box = QTextEdit(); self.log_box.setReadOnly(True); layout.addWidget(self.log_box)

    def _browse(self):
        path, _ = QFileDialog.getOpenFileName(self, "Choose workbook", "", "Excel workbook (*.xlsx);;All files (*.*)")
        if path: self.xlsx_ed.setText(path)

    def _start(self):
        path = Path(self.xlsx_ed.text().strip())
        if not path.is_file(): QMessageBox.warning(self, "Input required", "Choose an existing .xlsx workbook first."); return
        if not QMessageBox.question(self, "Update workbook", "The workbook will be updated in place. Continue?", QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes: return
        self.stop_event.clear(); self.log_box.clear(); self.start_btn.setEnabled(False); self.stop_btn.setEnabled(True)
        self.progress_bar.setRange(0, 1); self.progress_bar.setValue(0)
        self.worker = BillScraperWorker(path, self.delay.value(), self.stop_event)
        self.worker.progress.connect(self._progress); self.worker.done.connect(self._done); self.worker.failed.connect(self._failed); self.worker.start()

    def _stop(self): self.stop_event.set(); self.progress.setText("Stopping after the current bill...")
    def _progress(self, done, total, account, status):
        self.progress_bar.setRange(0, max(total, 1)); self.progress_bar.setValue(done)
        self.progress.setText(f"Processed {done} of {total} | {account}: {status}"); self.log_box.append(f"{account}: {status}")
    def _done(self, path, cancelled):
        self.start_btn.setEnabled(True); self.stop_btn.setEnabled(False); self.progress.setText("Stopped." if cancelled else "Completed")
        if not cancelled: notify(self, "Completed", f"Workbook updated:\n{path}", file_path=path)
    def _failed(self, message):
        self.start_btn.setEnabled(True); self.stop_btn.setEnabled(False); notify(self, "Error", message, critical=True)


class FileMoverWorker(QThread):
    done = Signal(int, list)
    failed = Signal(str)

    def __init__(self, excel, source, destination):
        super().__init__(); self.excel = excel; self.source = source; self.destination = destination

    def run(self):
        try:
            import pandas as pd
            df = pd.read_excel(self.excel, header=None)
            files = {}
            for name in os.listdir(self.source):
                path = os.path.join(self.source, name)
                if os.path.isfile(path): files[os.path.splitext(name)[0].strip()] = name
            moved = 0; missing = []
            for value in df.iloc[:, 0].dropna():
                sr = str(int(value)) if isinstance(value, (int, float)) and float(value).is_integer() else str(value).strip()
                if sr.endswith(".0"): sr = sr[:-2]
                if sr in files:
                    shutil.move(os.path.join(self.source, files[sr]), os.path.join(self.destination, files[sr])); moved += 1
                else: missing.append(sr)
            self.done.emit(moved, missing)
        except Exception as exc: self.failed.emit(str(exc))


class FileMoverPanel(QWidget):
    def __init__(self):
        super().__init__(); self.worker = None
        layout = QVBoxLayout(self); layout.setSpacing(10); layout.addWidget(lbl("PDF / File Mover", bold=True, color="#7eb8f7")); layout.addWidget(hline())
        group = QGroupBox("Inputs"); form = QVBoxLayout(group)
        self.excel_ed, self.source_ed, self.destination_ed = QLineEdit(), QLineEdit(), QLineEdit()
        for label, edit, kind in (("Excel file:", self.excel_ed, "excel"), ("Source folder:", self.source_ed, "source"), ("Destination folder:", self.destination_ed, "destination")):
            row = QHBoxLayout(); row.addWidget(QLabel(label)); edit.setReadOnly(True); edit.setProperty("preferred_width", 245); row.addWidget(edit, 1); button = QPushButton("Browse"); button.clicked.connect(lambda _, e=edit, k=kind: self._browse(e, k)); row.addWidget(button); form.addLayout(row)
        layout.addWidget(group)
        self.start_btn = QPushButton("▶  MOVE FILES"); self.start_btn.setStyleSheet(START_BTN); self.start_btn.clicked.connect(self._start); layout.addWidget(self.start_btn)
        self.status = QLabel("Status: Idle"); layout.addWidget(self.status); layout.addStretch()

    def _browse(self, edit, kind):
        path = QFileDialog.getOpenFileName(self, "Excel file", "", "Excel files (*.xlsx *.xls)")[0] if kind == "excel" else QFileDialog.getExistingDirectory(self, "Select folder")
        if path: edit.setText(path)

    def _start(self):
        if not all((self.excel_ed.text(), self.source_ed.text(), self.destination_ed.text())):
            QMessageBox.warning(self, "Missing input", "Select the Excel file, source folder, and destination folder."); return
        self.start_btn.setEnabled(False); self.status.setText("Status: Moving files...")
        self.worker = FileMoverWorker(self.excel_ed.text(), self.source_ed.text(), self.destination_ed.text())
        self.worker.done.connect(self._done); self.worker.failed.connect(self._failed); self.worker.finished.connect(lambda: self.start_btn.setEnabled(True)); self.worker.start()

    def _done(self, moved, missing):
        self.status.setText(f"Status: Completed — moved {moved} file(s)")
        message = f"Moved {moved} file(s).\n\n" + ("Missing files:\n\n" + "\n".join(missing) if missing else "All files moved successfully.")
        notify(self, "Completed", message, folder_path=self.destination_ed.text())

    def _failed(self, message): self.status.setText("Status: Failed"); notify(self, "Error", message, critical=True)


class CP52Worker(QThread):
    log = Signal(str)
    done = Signal(str)
    failed = Signal(str)

    def __init__(self, source_dir, workbook, output_dir, code_filter):
        super().__init__(); self.source_dir = Path(source_dir); self.workbook = Path(workbook); self.output_dir = Path(output_dir); self.code_filter = code_filter

    def run(self):
        try:
            # Reload during development so a long-running VS Code session
            # does not keep using an older CP-52 renderer after the source
            # file has been edited.
            import importlib
            from pitcsuite.tools import cp52_posting_checker_logic
            logic = importlib.reload(cp52_posting_checker_logic)
            result = logic.process_workbook(self.source_dir, self.workbook, self.output_dir, self.log.emit, self.code_filter)
            self.done.emit(str(result))
        except Exception as exc:
            self.failed.emit(str(exc))


class CP52PostingCheckerPanel(QWidget):
    def __init__(self):
        super().__init__(); self.worker = None
        layout = QVBoxLayout(self); layout.setSpacing(10)
        layout.addWidget(lbl("CP-52 Posting Checker", bold=True, color="#7eb8f7"))
        layout.addWidget(lbl("Match CP-52 posting records from TXT/PDF files and create highlighted PDFs.", color="#556")); layout.addWidget(hline())
        group = QGroupBox("Inputs")
        form = QVBoxLayout(group)
        self.source_ed, self.workbook_ed, self.output_ed = QLineEdit(), QLineEdit(), QLineEdit()
        for label, edit, kind in (("Posting folder:", self.source_ed, "folder"), ("Excel input:", self.workbook_ed, "excel"), ("Output folder:", self.output_ed, "folder")):
            row = QHBoxLayout(); row.addWidget(QLabel(label)); edit.setProperty("preferred_width", 245); row.addWidget(edit, 1)
            button = QPushButton("Browse"); button.clicked.connect(lambda _, e=edit, k=kind: self._browse(e, k)); row.addWidget(button); form.addLayout(row)
        filter_row = QHBoxLayout(); filter_row.addWidget(QLabel("Find Posting of:"))
        self.code_filter = QComboBox()
        self.code_filter.addItem("Code-C", "C")
        self.code_filter.addItem("Code-B", "B")
        self.code_filter.addItem("Code-F", "F")
        self.code_filter.addItem("All Codes", "ALL")
        self.code_filter.setCurrentIndex(3)
        filter_row.addWidget(self.code_filter); filter_row.addStretch(); form.addLayout(filter_row)
        layout.addWidget(group)
        actions = QHBoxLayout(); self.start_btn = QPushButton("▶  CHECK POSTING"); self.start_btn.setStyleSheet(START_BTN); self.start_btn.clicked.connect(self._start); actions.addWidget(self.start_btn)
        template_btn = QPushButton("Download Template"); template_btn.setMinimumWidth(145); template_btn.clicked.connect(self._download_template); actions.addWidget(template_btn)
        open_btn = QPushButton("Open Folder"); open_btn.setMinimumWidth(120); open_btn.clicked.connect(self._open_output); actions.addWidget(open_btn); actions.addStretch(); layout.addLayout(actions)
        self.status = QLabel("Status: Ready"); self.status.setStyleSheet("color:#4a90d9;"); layout.addWidget(self.status)
        self.log_box = QTextEdit(); self.log_box.setReadOnly(True); layout.addWidget(self.log_box)

    def _browse(self, edit, kind):
        if kind == "excel":
            path = QFileDialog.getOpenFileName(self, "Excel input", "", "Excel files (*.xlsx);;All files (*.*)")[0]
        else:
            path = QFileDialog.getExistingDirectory(self, "Select folder")
        if path: edit.setText(path)

    def _start(self):
        source, workbook = Path(self.source_ed.text().strip()), Path(self.workbook_ed.text().strip())
        if not source.is_dir() or not workbook.is_file():
            QMessageBox.warning(self, "Invalid input", "Select an existing posting folder and Excel workbook first."); return
        output = Path(self.output_ed.text().strip()) if self.output_ed.text().strip() else source / "output"
        self.output_ed.setText(str(output)); self.log_box.clear(); self.status.setText("Status: Processing..."); self.start_btn.setEnabled(False)
        self.worker = CP52Worker(source, workbook, output, self.code_filter.currentData())
        self.worker.log.connect(self._log); self.worker.done.connect(self._done); self.worker.failed.connect(self._failed); self.worker.finished.connect(lambda: self.start_btn.setEnabled(True)); self.worker.start()

    def _open_output(self):
        path = Path(self.output_ed.text().strip()) if self.output_ed.text().strip() else Path.cwd() / "output"
        path.mkdir(parents=True, exist_ok=True); os.startfile(str(path))

    def _download_template(self):
        download_template(
            self,
            ["AC No", "Amount", "Para Sr No.", "Debit Code", "Found Amount", "Remarks"],
            "CP52_Posting_Checker_Template.xlsx",
        )

    def _log(self, message):
        self.log_box.append(message)
        self.status.setText(f"Status: {message}")
    def _done(self, result):
        self.status.setText(f"Status: Completed — {result}")
        folder = self.output_ed.text().strip() or str(Path(self.source_ed.text()) / "output")
        notify(self, "Completed", f"Excel updated:\n{result}", file_path=result, folder_path=folder)
    def _failed(self, message):
        self.status.setText("Status: Failed"); notify(self, "Processing error", message, critical=True)


class CP136Worker(QThread):
    log = Signal(str)
    done = Signal(str, dict)
    failed = Signal(str)

    def __init__(self, file_paths, output_path, deduplicate, dedup_mode, only_present, include_status, include_filename):
        super().__init__()
        self.file_paths = file_paths
        self.output_path = output_path
        self.deduplicate = deduplicate
        self.dedup_mode = dedup_mode
        self.only_present = only_present
        self.include_status = include_status
        self.include_filename = include_filename

    def run(self):
        try:
            from pitcsuite.tools.cp136_posting_checker import run_cp136_check
            saved_path, stats = run_cp136_check(
                file_paths=self.file_paths,
                output_path=self.output_path,
                deduplicate=self.deduplicate,
                dedup_mode=self.dedup_mode,
                only_present=self.only_present,
                include_status=self.include_status,
                include_filename=self.include_filename,
                log_fn=self.log.emit
            )
            self.done.emit(saved_path, stats)
        except Exception as exc:
            self.failed.emit(str(exc))


class CP136PostingCheckerPanel(QWidget):
    def __init__(self):
        super().__init__()
        self.worker = None
        self.file_paths = []
        self.last_output = ""

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.addWidget(lbl("CP136 & 137 Posting Checker", bold=True, color="#7eb8f7"))
        layout.addWidget(lbl("Extract 14-figure Ref Nos, single postings, and field modifications from WAPDA RCP-136 & CP-137 reports into structured Excel.", color="#556"))
        layout.addWidget(hline())

        # Inputs Group
        group_in = QGroupBox("Source Reports (CP-136 & CP-137 .txt / .pdf files)")
        form_in = QVBoxLayout(group_in)
        self.file_list = QListWidget()
        self.file_list.setMaximumHeight(70)
        self.file_list.setMinimumHeight(50)
        form_in.addWidget(self.file_list)

        btn_row = QHBoxLayout()
        add_files_btn = QPushButton("📄 Add File(s)")
        add_files_btn.clicked.connect(self._add_files)
        btn_row.addWidget(add_files_btn)

        add_dir_btn = QPushButton("📁 Add Folder")
        add_dir_btn.clicked.connect(self._add_folder)
        btn_row.addWidget(add_dir_btn)

        remove_btn = QPushButton("✖ Remove Selected")
        remove_btn.clicked.connect(self._remove_selected)
        btn_row.addWidget(remove_btn)

        clear_btn = QPushButton("🗑 Clear All")
        clear_btn.clicked.connect(self._clear_files)
        btn_row.addWidget(clear_btn)
        btn_row.addStretch()
        form_in.addLayout(btn_row)
        layout.addWidget(group_in)

        # Output & Settings Group
        group_out = QGroupBox("Output Settings & Deduplication Options")
        form_out = QVBoxLayout(group_out)

        out_row = QHBoxLayout()
        out_row.addWidget(QLabel("Output destination:"))
        self.out_ed = QLineEdit()
        self.out_ed.setPlaceholderText("Auto-generated alongside source files if left blank")
        out_row.addWidget(self.out_ed, 3)

        browse_file_btn = QPushButton("📄 File...")
        browse_file_btn.clicked.connect(self._browse_file)
        out_row.addWidget(browse_file_btn)

        browse_folder_btn = QPushButton("📁 Folder...")
        browse_folder_btn.clicked.connect(self._browse_folder)
        out_row.addWidget(browse_folder_btn)
        out_row.addStretch(1)
        form_out.addLayout(out_row)

        # Strategy controls
        strategy_row = QHBoxLayout()
        strategy_row.addWidget(QLabel("Strategy:"))
        self.dedup_mode = QComboBox()
        self.dedup_mode.addItem("Prioritize POSTED (Recommended)", "prefer_posted")
        self.dedup_mode.addItem("Strictly Keep 1st Finding (Delete 2nd & subsequent)", "first")
        self.dedup_mode.addItem("Keep Last Finding", "last")
        strategy_row.addWidget(self.dedup_mode)
        strategy_row.addStretch()
        form_out.addLayout(strategy_row)

        # Extra options
        opts_row = QHBoxLayout()
        self.single_chk = QCheckBox("Single Postings only")
        self.single_chk.setChecked(True)
        self.single_chk.toggled.connect(self._on_dedup_toggled)
        opts_row.addWidget(self.single_chk)

        self.present_chk = QCheckBox("Only non-empty columns")
        self.present_chk.setChecked(True)
        opts_row.addWidget(self.present_chk)

        self.status_chk = QCheckBox("Include Status column")
        self.status_chk.setChecked(True)
        opts_row.addWidget(self.status_chk)

        self.file_chk = QCheckBox("Include Source File Name")
        opts_row.addWidget(self.file_chk)
        
        opts_row.addStretch()
        form_out.addLayout(opts_row)

        layout.addWidget(group_out)

        # Actions Row
        actions = QHBoxLayout()
        self.start_btn = QPushButton("▶  START PROCESSING")
        self.start_btn.setStyleSheet(START_BTN)
        self.start_btn.clicked.connect(self._start)
        actions.addWidget(self.start_btn)

        self.stop_btn = QPushButton("■  STOP")
        self.stop_btn.setStyleSheet(STOP_BTN)
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self._stop)
        actions.addWidget(self.stop_btn)

        self.open_excel_btn = QPushButton("📊 Open Excel")
        self.open_excel_btn.setEnabled(False)
        self.open_excel_btn.clicked.connect(self._open_excel)
        actions.addWidget(self.open_excel_btn)

        self.open_folder_btn = QPushButton("📂 Open Folder")
        self.open_folder_btn.setEnabled(False)
        self.open_folder_btn.clicked.connect(self._open_folder)
        actions.addWidget(self.open_folder_btn)
        actions.addStretch()
        layout.addLayout(actions)

        self.status = QLabel("Status: Ready. Add CP-136 / CP-137 file(s) or folder to begin.")
        self.status.setStyleSheet("color:#4a90d9;")
        layout.addWidget(self.status)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        layout.addWidget(self.log_box)

    def _on_dedup_toggled(self, checked):
        self.dedup_mode.setEnabled(checked)

    def _add_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Select CP-136 / CP-137 Reports", "", "Supported Reports (*.txt *.pdf);;Text Files (*.txt);;PDF Files (*.pdf);;All files (*.*)")
        for p in paths:
            norm_p = os.path.normpath(p)
            if norm_p not in self.file_paths:
                self.file_paths.append(norm_p)
                self.file_list.addItem(f"📄 {os.path.basename(norm_p)}  ({norm_p})")
        self._update_default_output()

    def _add_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder Containing CP-136 / CP-137 Files")
        if folder:
            import glob
            files = sorted(glob.glob(os.path.join(folder, "*.txt")) + glob.glob(os.path.join(folder, "*.pdf")))
            for tf in files:
                norm_p = os.path.normpath(tf)
                if norm_p not in self.file_paths:
                    self.file_paths.append(norm_p)
                    self.file_list.addItem(f"📄 {os.path.basename(norm_p)}  ({norm_p})")
            self._update_default_output()

    def _remove_selected(self):
        for item in self.file_list.selectedItems():
            row = self.file_list.row(item)
            self.file_list.takeItem(row)
            if row < len(self.file_paths):
                del self.file_paths[row]
        self._update_default_output()

    def _clear_files(self):
        self.file_paths.clear()
        self.file_list.clear()
        self.out_ed.clear()
        self.status.setText("Status: Cleared all files.")

    def _update_default_output(self):
        if self.file_paths and not self.out_ed.text().strip():
            first_dir = os.path.dirname(self.file_paths[0])
            first_base = os.path.splitext(os.path.basename(self.file_paths[0]))[0]
            suffix = "_Single_Postings.xlsx" if self.single_chk.isChecked() else "_Extracted.xlsx"
            self.out_ed.setText(os.path.normpath(os.path.join(first_dir, f"{first_base}{suffix}")))
        if len(self.file_paths) > 1:
            self.file_chk.setChecked(True)

    def _browse_file(self):
        curr = self.out_ed.text().strip()
        init_dir = os.path.dirname(curr) if curr and os.path.isdir(os.path.dirname(curr)) else ""
        path, _ = QFileDialog.getSaveFileName(self, "Save Output Excel As", init_dir, "Excel Workbook (*.xlsx)")
        if path:
            if not path.lower().endswith('.xlsx'):
                path += '.xlsx'
            self.out_ed.setText(os.path.normpath(path))

    def _browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder:
            curr = self.out_ed.text().strip()
            name = os.path.basename(curr) if curr and curr.lower().endswith('.xlsx') else "CP136_137_Posting_Results.xlsx"
            self.out_ed.setText(os.path.normpath(os.path.join(folder, name)))

    def _start(self):
        if not self.file_paths:
            QMessageBox.warning(self, "No Files", "Please add at least one CP-136 or CP-137 report file (.txt or .pdf).")
            return

        out_path = self.out_ed.text().strip()
        self.log_box.clear()
        self.status.setText("Status: Processing CP-136 / CP-137 reports...")
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)

        self.worker = CP136Worker(
            file_paths=list(self.file_paths),
            output_path=out_path,
            deduplicate=self.single_chk.isChecked(),
            dedup_mode=self.dedup_mode.currentData(),
            only_present=self.present_chk.isChecked(),
            include_status=self.status_chk.isChecked(),
            include_filename=self.file_chk.isChecked()
        )
        self.worker.log.connect(self._log)
        self.worker.done.connect(self._done)
        self.worker.failed.connect(self._failed)
        self.worker.finished.connect(lambda: (self.start_btn.setEnabled(True), self.stop_btn.setEnabled(False)))
        self.worker.start()

    def _stop(self):
        if self.worker and self.worker.isRunning():
            self.worker.requestInterruption()
            self.status.setText("Status: Stopping worker...")

    def _log(self, msg):
        self.log_box.append(msg)

    def _done(self, saved_path, stats):
        self.last_output = saved_path
        self.open_excel_btn.setEnabled(True)
        self.open_folder_btn.setEnabled(True)

        total = stats.get('total_records', 0)
        dups = stats.get('duplicates_removed', 0)
        posted = stats.get('posted_records', 0)
        unposted = stats.get('not_posted_records', 0)

        dup_msg = f" ({dups} duplicates removed)" if dups > 0 else ""
        self.status.setText(f"Status: Completed — {total} single postings exported{dup_msg}")

        msg_body = (
            f"Successfully processed {stats.get('total_files', 1)} file(s)!\n\n"
            f"• Single Postings: {total}\n"
            f"• Duplicates Removed: {dups}\n"
            f"• Posted: {posted}\n"
            f"• Not Posted: {unposted}\n\n"
            f"Excel file created at:\n{saved_path}"
        )
        notify(self, "CP136 & 137 Extraction Complete", msg_body, file_path=saved_path, folder_path=os.path.dirname(saved_path))

    def _failed(self, error_msg):
        self.status.setText("Status: Failed")
        notify(self, "CP136 & 137 Error", error_msg, critical=True)

    def _open_excel(self):
        if self.last_output and os.path.exists(self.last_output):
            os.startfile(self.last_output)

    def _open_folder(self):
        if self.last_output:
            folder = os.path.dirname(os.path.abspath(self.last_output))
            if os.path.exists(folder):
                os.startfile(folder)

class CP21cWorker(QThread):
    log = Signal(str)
    progress = Signal(int, int, str)
    done = Signal(str)
    failed = Signal(str)

    def __init__(self, pdf_path, excel_path):
        super().__init__()
        self.pdf_path = pdf_path
        self.excel_path = excel_path

    def run(self):
        try:
            from pitcsuite.tools.cp21c_converter import convert_cp21c_pdf
            
            def progress_cb(current, total, msg):
                self.progress.emit(current, total, msg)

            stats = convert_cp21c_pdf(
                self.pdf_path, 
                self.excel_path, 
                progress_callback=progress_cb, 
                log_callback=self.log.emit
            )
            self.done.emit(self.excel_path)
        except Exception as exc:
            self.failed.emit(str(exc))

class CP21cConverterPanel(QWidget):
    def __init__(self):
        super().__init__()
        self.worker = None

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.addWidget(lbl("CP-21C Converter", bold=True, color="#7eb8f7"))
        layout.addWidget(lbl("Extract CP-21C electricity meter reading data from PDF to structured Excel.", color="#556"))
        layout.addWidget(hline())

        group = QGroupBox("Inputs")
        form = QVBoxLayout(group)
        
        self.pdf_ed = QLineEdit()
        self.pdf_ed.setReadOnly(True)
        r1 = QHBoxLayout()
        r1.addWidget(QLabel("Source CP-21C PDF:"))
        r1.addWidget(self.pdf_ed, 1)
        b1 = QPushButton("Browse")
        b1.clicked.connect(self._browse_pdf)
        r1.addWidget(b1)
        form.addLayout(r1)
        
        self.excel_ed = QLineEdit()
        r2 = QHBoxLayout()
        r2.addWidget(QLabel("Target Excel File:"))
        r2.addWidget(self.excel_ed, 1)
        b2 = QPushButton("Browse")
        b2.clicked.connect(self._browse_excel)
        r2.addWidget(b2)
        form.addLayout(r2)
        
        layout.addWidget(group)

        actions = QHBoxLayout()
        self.start_btn = QPushButton("▶  START CONVERSION")
        self.start_btn.setStyleSheet(START_BTN)
        self.start_btn.clicked.connect(self._start)
        actions.addWidget(self.start_btn)
        
        actions.addStretch()
        layout.addLayout(actions)

        self.status = QLabel("Status: Ready")
        self.status.setStyleSheet("color:#4a90d9;")
        layout.addWidget(self.status)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        layout.addWidget(self.log_box)

    def _browse_pdf(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select CP-21C PDF", "", "PDF Files (*.pdf);;All files (*.*)")
        if path:
            self.pdf_ed.setText(path)
            if not self.excel_ed.text():
                base = os.path.splitext(path)[0]
                self.excel_ed.setText(f"{base}.xlsx")

    def _browse_excel(self):
        path, _ = QFileDialog.getSaveFileName(self, "Select Excel Output", "", "Excel Workbook (*.xlsx)")
        if path:
            if not path.lower().endswith('.xlsx'):
                path += '.xlsx'
            self.excel_ed.setText(path)

    def _start(self):
        pdf_path = self.pdf_ed.text().strip()
        excel_path = self.excel_ed.text().strip()
        
        if not pdf_path or not os.path.exists(pdf_path):
            QMessageBox.warning(self, "Missing input", "Select a valid CP-21C PDF file.")
            return
        if not excel_path:
            QMessageBox.warning(self, "Missing input", "Specify a target Excel file.")
            return

        self.log_box.clear()
        self.status.setText("Status: Converting...")
        self.progress_bar.setValue(0)
        self.start_btn.setEnabled(False)

        self.worker = CP21cWorker(pdf_path, excel_path)
        self.worker.log.connect(self.log_box.append)
        self.worker.progress.connect(self._progress)
        self.worker.done.connect(self._done)
        self.worker.failed.connect(self._failed)
        self.worker.start()

    def _progress(self, current, total, msg):
        pct = int((current / total) * 100)
        self.progress_bar.setValue(pct)
        self.status.setText(f"Status: {msg}")

    def _done(self, excel_path):
        self.start_btn.setEnabled(True)
        self.progress_bar.setValue(100)
        self.status.setText("Status: Completed")
        notify(self, "Completed", "CP-21C conversion finished successfully.", file_path=excel_path)

    def _failed(self, msg):
        self.start_btn.setEnabled(True)
        self.status.setText("Status: Failed")
        notify(self, "Error", msg, critical=True)

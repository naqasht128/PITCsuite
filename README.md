# PITC Suite

[![Python Version](https://img.shields.io/badge/python-3.10%20%7C%203.11%20%7C%203.12%20%7C%203.14-blue.svg)](https://www.python.org/)
[![GUI Framework](https://img.shields.io/badge/GUI-PySide6%20(Qt6)-41cd52.svg)](https://wiki.qt.io/Qt_for_Python)
[![Platform](https://img.shields.io/badge/platform-Windows-0078d6.svg)](https://www.microsoft.com/windows)
[![Version](https://img.shields.io/badge/version-2.8.91-orange.svg)](#)

**PITC Suite** is an all-in-one desktop productivity and automation suite designed for utility billing operations, power sector reconciliation, and batch document processing. Built on **PySide6 (Qt 6)**, it combines automated scraping, 2FA-secured portal downloads, ledger audit checkers, and industrial-grade PDF manipulation tools into a unified, responsive interface.

---

## 🌟 Key Highlights

- **16 Integrated Tools:** Complete workflow coverage from automated portal ingestion to billing reconciliation and batch printing, now including Government Assessment & Realization extraction and CP 139 posting verification.
- **Polished UX/UI:** Crisp SVG icons, ergonomic file/folder opening directly from completion dialogs, and a responsive PySide6 dark theme.
- **Fusion Dark Design System:** Clean dark theme with ergonomic color hierarchy (vibrant green primary actions, orange secondary actions, and high-legibility typography).
- **Collapsible Navigation Drawer:** Fast switching between full label mode and an icon-only compact sidebar.
- **Asynchronous Processing:** Long-running I/O tasks, web scraping, and PDF rendering execute on background threads (`QThread`) with live progress reporting and cancellation support.
- **Standalone Distribution:** Pre-configured with PyInstaller specifications to generate single-file Windows executables.

---

## 🛠️ Integrated Tools

| Tool | Category | Description |
| :--- | :--- | :--- |
| **📥 PITC Report Downloader** | Web Automation | Automated document and report retrieval from the PITC portal with automated session management and TOTP/2FA generation via pyotp. |
| **🔍 Extract Pages from PITC Reports** | PDF Processing | High-speed targeted page extraction, splitting, and section filtering from large PITC billing documents. |
| **🧾 E/Bills Downloader** | Web Automation | High-throughput batch utility bill downloader with automated queuing and retry logic. |
| **🧮 Bill Scrapper** | Data Extraction | Scrapes and extracts billing metadata and consumer records directly into structured formats. |
| **💰 Payment Extractor from 88L** | Data Extraction | Parses and extracts payment scrolls and disbursement figures from 88L report sheets. |
| **📋 CP 52 Posting to PDF** | Reconciliation | Automated verification and audit reconciliation for CP-52 cash collection and posting registers into highlighted PDFs. |
| **📝 CP 136 & 137 Posting Checker** | Reconciliation | Parses WAPDA RCP-136 & CP-137 change posting reports (.txt and .pdf), extracts single postings, deletes duplicate findings, and exports multi-sheet structured Excel workbooks. |
| **📑 CP 139 Posting Checker** | Reconciliation | Parses WAPDA RCP-139 Unidentified Cash / Debtor's Adjustment posting reports (.txt and .pdf), extracts Ref#, S/Div, Grand Total, and posting status with error codes into styled Excel spreadsheets. |
| **🧾 CP 22T Posting Checker** | Reconciliation | Verifies and validates CP22T revenue posting transactions against reference records. |
| **⚡ CP21c Converter** | Data Extraction | Extracts CP-21C electricity meter reading data from PDF and exports it to a structured single-sheet Excel file. |
| **🔁 Double CR Checker** | Reconciliation | Audits consumer ledger records to detect duplicate consumer references (CR) and redundant postings. |
| **📊 Govt Assessment Realization** | Data Extraction | Parses WAPDA Assessment & Realization reports for Government Departments (.txt), extracts S/Div, 14-digit Ref#, Status, Name & Address, Tariff, Department Code, Arrear Age, OP. Balance, Payments Normal, Adjustor, Write Off, and Net Arrear into styled Excel spreadsheets. |
| **📑 PDF Merger** | PDF Processing | Multi-folder merging (up to 5 folders simultaneously), natural numerical ordering, duplex/N-up blank-page padding, and filename header/footer stamping. |
| **📂 File Checker** | Data Audit | Scans directories to identify missing serial numbers, sequence gaps, or malformed filenames in batch outputs. |
| **📦 File Mover** | File Operations | Categorizes, sorts, and organizes batch generated files according to configurable pattern rules. |
| **🏷️ File Renamer** | File Operations | Bulk renaming utility designed for billing cycles, batch codes, and document standardisation. |
| **⚙️ Settings** | Configuration | Centralized management of credentials (username, password, OTP secrets) and local storage directories. |

---

## 🏗️ Architecture & Tech Stack

```text
PITC Suite
├── Pitcsuite.py              # Entrypoint launcher
├── config.example.ini        # Credential & configuration template
├── pitcsuite.spec            # PyInstaller build specification
├── pitcsuite/
│   ├── app.py                # Application bootstrap & lifecycle
│   ├── main_window.py        # Main window, navigation drawer & stack management
│   ├── config.py             # Runtime config resolver
│   ├── icons.py              # Asset loaders & icon helpers
│   ├── ui_helpers.py         # Stylesheet definitions, palettes & helper widgets
│   ├── assets/               # Application icons and branding
│   └── tools/                # Modular tool implementations
│       ├── assessment_realization.py
│       ├── bill_downloader.py
│       ├── bill_scraper_logic.py
│       ├── cp136_posting_checker.py
│       ├── cp139_posting_checker.py
│       ├── cp21c_converter.py
│       ├── cp22t_posting_checker.py
│       ├── cp52_posting_checker_logic.py
│       ├── double_cr.py
│       ├── extract_page.py
│       ├── file_checker.py
│       ├── file_renamer.py
│       ├── integrated_tools.py
│       ├── payment_extract_88l.py
│       ├── pdf_merger.py
│       ├── pitc_downloader.py
│       └── settings.py
└── tests/                    # Automated test suite
    └── test_pdf_merger.py
```

- **Frontend / GUI:** PySide6 (Qt 6)
- **Document & PDF Manipulation:** `pypdf`, `reportlab`, `PyMuPDF` (`fitz`), `pdfplumber`
- **Automation & Scraping:** `selenium`, `webdriver-manager`, `beautifulsoup4`, `requests`
- **Data Analysis & Processing:** `pandas`, `openpyxl`
- **Security & 2FA:** `pyotp`

---

## 🚀 Getting Started

### Prerequisites

- **Python 3.10+** (Tested on Python 3.10, 3.11, 3.12, 3.14)
- **Google Chrome** (required for automated browser-based scrapers and downloaders)
- **Git**

### Installation

1. **Clone the repository:**
   ```bash
   git clone https://bitbucket.org/tnaqash/pitcsuite.git
   cd PitcSuite
   ```

2. **Create and activate a virtual environment (recommended):**
   ```bash
   # Windows PowerShell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

---

## ⚙️ Configuration

1. Copy the example configuration file:
   ```bash
   copy config.example.ini config.ini
   ```

2. Configure your credentials in `config.ini`:
   ```ini
   [credentials]
   username = your_portal_username
   password = your_portal_password
   otp_secret = your_totp_secret_key
   ```
   > [!NOTE]
   > You can also update credentials directly within the application via the **⚙️ Settings** panel.
   > Real credentials in `config.ini` are strictly ignored by `.gitignore` to prevent accidental credential leakage.

---

## 🖥️ Running the Application

Launch the application directly with Python:

```bash
python Pitcsuite.py
```

---

## 📦 Building Standalone Executable

The application includes a fully configured [pitcsuite.spec](pitcsuite.spec) file bundling all assets, icons, and dependencies.

To build the standalone Windows executable:

```bash
python -m PyInstaller --noconfirm pitcsuite.spec
```

The compiled binary will be placed in the `dist/` directory:
```text
dist/
└── pitcsuite.exe
```

---

## 🧪 Testing

Run the automated unit test suite:

```bash
python -m unittest discover -s tests -v
```

---

## 🔒 Security Best Practices

- **Never commit `config.ini`:** Always use `config.example.ini` for version control. Real credentials, passwords, and 2FA secrets must remain local.
- **Push Protection:** Ensure sensitive tokens and API keys are not hardcoded into any Python scripts or test cases.

---

## 👤 Author

Developed by **Tahir Naqash**


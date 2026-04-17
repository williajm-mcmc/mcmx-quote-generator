# MCMX Quote Generator — v2.0

A desktop application for building and exporting professional project quotes as formatted Word documents (`.docx`). Built with PyQt6. No Python installation required to run.

---

## What's New in v2.0

### Bug Fixes
- **Special-character crash** — project names, customer names, and other text fields containing `&`, `<`, `>`, or `"` (e.g. "AT&T") no longer crash the document generator with an XML parse error.
- **Section reorder arrows** — the ▲ / ▼ buttons on custom sections now correctly reorder sections in the tool and in the generated document.
- **Insert Table dialog** — the row/column spinners now show their up/down arrows; the "Add a Total row" and "Columns to sum" checkboxes now show their indicators.
- **Proposal number on cover page** — the proposal number printed on the document's first page now matches the saved filename exactly, including the date stamp and version (e.g. `P-2024-001-20240416 V1.0`).
- **Table of Contents page numbers** — ToC entries no longer all show page 1. Fixed a `settings.xml` regex that could leave conflicting `updateFields` elements, and updated the TOC field instruction to use `\o "1-9" \u` (matching what Word's ribbon generates) for reliable page-number updates across Word versions.

### Quote Generator — BOM / Service Rows
- **MCMX-SERVICES-EXPENSES bold** — the service name is now **bold** in the generated Word document.
- **Reference-only breakdown** — drive time (TRAVEL) and the mileage / meals / hotel breakdown (EXPENSES) remain visible in the BOM tool for reference but are automatically stripped from the generated Word document. A note below the service-row buttons reminds you of this.
- **Expenses mileage bullet** — the mileage bullet now reads `• Mileage` without a count, for a cleaner document appearance.
- **Service row font sizes** — subtext in all three service rows (AFSE / TRAVEL / EXPENSES) renders at **8 pt**; the AFSE "Rates:" label renders at **9 pt bold**, matching the intended document design.

### IBE & Thermal Estimators
- **Travel In / Travel Out** — these schedule rows are now suppressed when no overnight stay is required (all techs driving under 2 hours and not flying). Local crews see only their per-day work rows.

---

## Features

| Area | Capability |
|------|-----------|
| **Quote Generator** | Project header (name, customer, location, contact, proposal number), customer site photo (drag-drop), versioned save/load (`.mcmxq`), generate `.docx` |
| **BOM Tables** | Part number, qty, price, margin, line total; hours toggle; global margin control; pre-filled MCMX service rows (AFSE, Travel, Expenses); rich cell editor with bullets and tables; custom section headers |
| **Paragraph Sections** | Rich-text editor — bold, italic, headers, nested bullets, inline tables |
| **Table of Contents** | Live side-panel ToC; injected into generated document with correct page numbers |
| **Version Tracking** | Major/minor version spinners; change-history log saved with project |
| **OTTO Tool** | OTTO-specific quoting with its own Word template |
| **Thermal Imaging** | Multi-tech scheduling, overnight/travel logic, per-tech travel mode |
| **IBE Estimator** | IBE-specific schedule builder with conditional travel/hotel rows |
| **Cost Estimator** | Excel-style cost breakdown; export/import `.mcmxc`; open in Excel |

---

## Download & Install

Pre-built releases for Windows and macOS are on the **[Releases](../../releases)** page. No Python required.

### Windows — Fresh Install
1. Download `QuoteGenerator-Windows.zip` from the [latest release](../../releases/latest)
2. Extract the folder anywhere (Desktop, `C:\Program Files\`, etc.)
3. Double-click `QuoteGenerator.exe`

### macOS — Fresh Install
1. Download `QuoteGenerator-Mac.zip` from the [latest release](../../releases/latest)
2. Extract and drag `QuoteGenerator.app` to your **Applications** folder
3. Double-click to open
4. If macOS blocks it: right-click → **Open** → **Open**

---

## Upgrading from v1.1 → v2.0

There is no system installer — upgrading is just a folder swap. Your `.mcmxq` project files are stored separately and are fully compatible with v2.0; no migration needed.

### Windows
1. **Remove v1.1** — delete the old `QuoteGenerator` folder (the one containing `QuoteGenerator.exe`).
2. **Install v2.0** — extract `QuoteGenerator-Windows.zip` and run `QuoteGenerator.exe`.

### macOS
1. **Remove v1.1** — drag `QuoteGenerator.app` from Applications to the Trash and empty it.
2. **Install v2.0** — extract `QuoteGenerator-Mac.zip` and drag `QuoteGenerator.app` to Applications.

---

## Building from Source

### Prerequisites

```bash
pip install PyQt6 python-docx docxtpl pillow pyinstaller
```

> `markupsafe`, `jinja2`, and `lxml` are pulled in automatically as dependencies of `docxtpl`.

### Windows

```bat
build_windows.bat
```

Output: `dist\QuoteGenerator\QuoteGenerator.exe` — zip the `dist\QuoteGenerator` folder to distribute.

### macOS

```bash
bash build_mac.sh
```

Output: `dist/QuoteGenerator.app` — run `cd dist && zip -r QuoteGenerator-Mac.zip QuoteGenerator.app` to distribute.

---

## Project File Format

Projects are saved as `.mcmxq` files (JSON). They store all form fields, BOM tables, rich-text content, version history, and estimator data. Open via **File → Import MCMXQ**.

Cost Estimator data is also exported as `.mcmxc` (JSON), which can be opened directly in Excel via the Cost Estimator tab.

---

## CI / Releases

Pushing a version tag triggers the GitHub Actions workflow, which builds both platforms, optionally code-signs the Windows executable, and publishes a release with both zip attachments:

```bash
git tag v2.0
git push origin v2.0
```

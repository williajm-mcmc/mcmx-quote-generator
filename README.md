# MCMX Quote Generator

A desktop application for building and exporting professional project quotes as formatted Word documents (`.docx`). Built with PyQt6.

---

## Features

- **Project header fields** — project name, customer name, location, contact info, and proposal number
- **Customer picture** — drag-and-drop or browse to attach a customer site photo
- **Bill of Material sections** — add multiple BOM tables with part numbers, quantities, pricing, and automatic margin and total calculations
- **Paragraph sections** — rich-text editor supporting bold, italic, headers, nested bullet lists, and inline tables
- **Live Table of Contents** — auto-updating TOC panel reflecting all sections and sub-headers
- **Version tracking** — major/minor version spinners and a change history log, both saved with the project
- **Save / Load projects** — saves all content to a `.mcmxq` file for later editing; filename includes the proposal number, date, and version (e.g. `MCMX-BOIL-FTSC-20260325 V1.0.mcmxq`)
- **Generate Quote** — exports a fully formatted `.docx` Word document using the same naming convention

---

## Download

Pre-built releases for Windows and macOS are available on the [Releases](../../releases) page. No Python installation required.

### Windows
1. Download `QuoteGenerator-Windows.zip`
2. Extract the folder anywhere
3. Double-click `QuoteGenerator.exe`

### macOS
1. Download `QuoteGenerator-Mac.zip`
2. Extract and drag `QuoteGenerator.app` to your Applications folder
3. Double-click to open
4. If macOS shows an "unidentified developer" warning: right-click the app → **Open** → **Open**

---

## Building from Source

### Prerequisites

```
Python 3.12+
pip install PyQt6 python-docx docxtpl pyinstaller pillow
```

### Windows

```
build_windows.bat
```

Output: `dist\QuoteGenerator\QuoteGenerator.exe`

### macOS

```
bash build_mac.sh
```

Output: `dist/QuoteGenerator.app`

---

## Project File Format

Projects are saved as `.mcmxq` files (JSON). They store all form fields, BOM tables, paragraph content, version info, and change history. Import a saved project via **File → Import MCMXQ**.

---

## Releases / CI

Builds are automated via GitHub Actions. Pushing a version tag triggers both the Windows and macOS builds and publishes a release automatically:

```bash
git tag v1.1
git push origin v1.1
```

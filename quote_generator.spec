# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller spec for MCMX Quote Generator
# Build on Windows  -> produces dist/QuoteGenerator/QuoteGenerator.exe
# Build on macOS    -> produces dist/QuoteGenerator.app
#
# Usage:
#   Windows:  pyinstaller quote_generator.spec
#   macOS:    pyinstaller quote_generator.spec

import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# ── Data files to bundle ──────────────────────────────────────────────────────
added_datas = [
    ('logo.png',              '.'),
    ('template.docx',         '.'),
    ('otto_template.docx',    '.'),
    ('thermal_template.docx', '.'),
    ('mainwindow.ui',         '.'),
]

# docxtpl / jinja2 templates
added_datas += collect_data_files('docxtpl')
added_datas += collect_data_files('jinja2')

# ── Hidden imports ────────────────────────────────────────────────────────────
hidden = [
    'updater',                # auto-update helper (imported inside method bodies)
    'PyQt6.QtPrintSupport',   # required by some Qt widgets at runtime
    'docx',
    'docx.oxml',
    'docx.oxml.ns',
    'docxtpl',
    'jinja2',
    'lxml',
    'lxml.etree',
    'lxml._elementpath',
    'html.parser',
    'openpyxl',               # IBE / Cost Estimator Excel export
    'openpyxl.styles',
    'openpyxl.utils',
]
hidden += collect_submodules('docx')
hidden += collect_submodules('docxtpl')
hidden += collect_submodules('openpyxl')

# ── Unused modules to exclude ─────────────────────────────────────────────────
# Stripping unused PyQt6 components and stdlib bloat saves ~15-20 MB.
excluded = [
    # Unused PyQt6 modules (each pulls in large Qt DLLs / frameworks)
    'PyQt6.Qt3DAnimation', 'PyQt6.Qt3DCore', 'PyQt6.Qt3DExtras',
    'PyQt6.Qt3DInput',     'PyQt6.Qt3DLogic', 'PyQt6.Qt3DRender',
    'PyQt6.QtBluetooth',   'PyQt6.QtCharts',  'PyQt6.QtDataVisualization',
    'PyQt6.QtDesigner',    'PyQt6.QtHelp',    'PyQt6.QtLocation',
    'PyQt6.QtMultimedia',  'PyQt6.QtMultimediaWidgets',
    'PyQt6.QtNfc',         'PyQt6.QtPositioning',
    'PyQt6.QtQuick',       'PyQt6.QtQuick3D',
    'PyQt6.QtRemoteObjects','PyQt6.QtSensors',
    'PyQt6.QtSerialBus',   'PyQt6.QtSerialPort',
    'PyQt6.QtSql',         'PyQt6.QtSvg', 'PyQt6.QtSvgWidgets',
    'PyQt6.QtTest',        'PyQt6.QtWebChannel',
    'PyQt6.QtWebEngineCore','PyQt6.QtWebEngineQuick',
    'PyQt6.QtWebEngineWidgets','PyQt6.QtWebSockets',
    'PyQt6.QtXml',
    # Unused stdlib / large packages
    'tkinter', '_tkinter',
    'unittest', 'test',
    'distutils', 'setuptools', 'pip',
    'sqlite3', '_sqlite3',
    'numpy', 'scipy', 'matplotlib', 'pandas',
    'cryptography', 'OpenSSL',
]

a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=added_datas,
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excluded,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# ── Windows / Linux: single-folder EXE ───────────────────────────────────────
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='QuoteGenerator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,                      # UPX corrupts python3xx.dll — keep off
    upx_exclude=[],
    console=False,                 # no console window
    icon='logo.ico',               # Windows requires .ico
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,                      # UPX corrupts python3xx.dll — keep off
    upx_exclude=[],
    name='QuoteGenerator',
)

# ── macOS: wrap the folder in a .app bundle ───────────────────────────────────
if sys.platform == 'darwin':
    app = BUNDLE(
        coll,
        name='QuoteGenerator.app',
        icon='logo.png',           # must be .icns on macOS; see note below
        bundle_identifier='com.mcmx.quotegenerator',
        info_plist={
            'NSHighResolutionCapable': True,
            'CFBundleShortVersionString': '2.1',
            'CFBundleVersion': '2.1',
            'NSRequiresAquaSystemAppearance': False,  # allow dark mode
        },
    )

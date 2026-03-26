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
    ('logo.png',       '.'),
    ('template.docx',  '.'),
    ('mainwindow.ui',  '.'),
]

# docxtpl / jinja2 templates
added_datas += collect_data_files('docxtpl')
added_datas += collect_data_files('jinja2')

# ── Hidden imports ────────────────────────────────────────────────────────────
hidden = [
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
    # 'PIL',        # only needed if app uses Pillow directly
]
hidden += collect_submodules('docx')
hidden += collect_submodules('docxtpl')

a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=added_datas,
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    upx=True,
    console=False,                 # no console window
    icon='logo.ico',               # Windows requires .ico
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
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
            'CFBundleShortVersionString': '1.0',
            'CFBundleVersion': '1.0',
            'NSRequiresAquaSystemAppearance': False,  # allow dark mode
        },
    )

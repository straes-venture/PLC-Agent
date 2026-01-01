# -*- mode: python ; coding: utf-8 -*-

import sys
from pathlib import Path

# PyInstaller runs the spec from the project root
project_root = Path.cwd()

block_cipher = None


a = Analysis(
    ['ui.py'],
    pathex=[str(project_root)],
    binaries=[],
    datas=[
        # --- REQUIRED NON-PYTHON ASSETS ---
        (str(project_root / 'rss' / 'Auto_Parse.RO0'), 'rss'),

        # Bundle folders to avoid first-run edge cases
        (str(project_root / 'rss'), 'rss'),
        (str(project_root / 'reports'), 'reports'),
        (str(project_root / 'exports'), 'exports'),
    ],
    hiddenimports=[
        # pywinauto
        'pywinauto',
        'pywinauto.application',
        'pywinauto.keyboard',
        'pywinauto.findwindows',

        # PDF extraction
        'fitz',              # PyMuPDF
        'pdfplumber',
        'pdfminer',
        'pdfminer.high_level',

        # Windows / printing / clipboard
        'win32print',
        'win32clipboard',
        'win32con',

        # Your modules
        'python.open_program',
        'python.print_report',
        'python.close_apps',
        'extract.dump_ladders',
        'digest.program_snapshot',
        'util.common',
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)


pyz = PYZ(
    a.pure,
    a.zipped_data,
    cipher=block_cipher
)


exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='PLC_Agent',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,        # NEVER use UPX with pywinauto
    console=False,    # UI app
)


coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    name='PLC_Agent',
)

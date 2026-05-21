# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for Sterling B2Bi Map Migration Accelerator
This creates a standalone desktop application bundle
"""

import sys
from pathlib import Path

block_cipher = None

# Define the main script
main_script = 'tkinter_app.py'

# All Python modules that need to be included
hiddenimports = [
    'tkinter',
    'tkinter.ttk',
    'tkinter.filedialog',
    'tkinter.messagebox',
    'tkinter.scrolledtext',
    'pandas',
    'openpyxl',
    'openpyxl.styles',
    'openpyxl.utils',
    'openpyxl.worksheet',
    'openpyxl.workbook',
    'xml.etree.ElementTree',
    'xml.dom.minidom',
    'zipfile',
    'shutil',
    'pathlib',
    'datetime',
    'logging',
    'threading',
    'queue',
    'subprocess',
    're',
    'csv',
    'json',
]

# Data files to include (Excel templates, etc.)
datas = [
    ('Generic_checklistMain.xlsm', '.'),
    ('mapping_results.xlsx', '.'),
    ('process_data_rules.xlsx', '.'),
    ('README.md', '.'),
]

# Python scripts that need to be bundled
scripts_to_bundle = [
    'backend_api.py',
    'mxl_parser.py',
    'mxl_processor.py',
    'codelist_extractor.py',
    'codelist_renamer.py',
    'error_reporter.py',
    'process_data_updater_v3.py',
    'process_data_manager.py',
    'inbound_mapsFeatures.py',
    'outbound_mapsFeatures.py',
    'modify_character_encoding.py',
    'modify_syntax_token.py',
    'remove_namespace_prefixes.py',
    'update_presession_rules.py',
    'process_all_mxl_files.py',
]

# Add scripts as data files so they can be called as subprocesses
for script in scripts_to_bundle:
    datas.append((script, '.'))

a = Analysis(
    [main_script],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'numpy',
        'scipy',
        'PIL',
        'PyQt5',
        'PyQt6',
        'PySide2',
        'PySide6',
        'wx',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='SterlingMapMigrator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # Set to False for GUI app (no console window)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Add icon file path here if you have one
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='SterlingMapMigrator',
)

# For macOS, create an app bundle
if sys.platform == 'darwin':
    app = BUNDLE(
        coll,
        name='SterlingMapMigrator.app',
        icon=None,  # Add icon file path here if you have one
        bundle_identifier='com.sterling.mapmigrator',
        info_plist={
            'NSPrincipalClass': 'NSApplication',
            'NSHighResolutionCapable': 'True',
            'CFBundleShortVersionString': '1.0.0',
            'CFBundleVersion': '1.0.0',
            'CFBundleName': 'Sterling Map Migrator',
            'CFBundleDisplayName': 'Sterling B2Bi Map Migration Accelerator',
        },
    )
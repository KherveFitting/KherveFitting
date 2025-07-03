# -*- mode: python ; coding: utf-8 -*-
import os

# Get current directory
current_dir = os.path.dirname(os.path.abspath('KherveDB.spec'))

a = Analysis(
    ['libraries/LibraryID.py'],
    pathex=[current_dir],
    binaries=[],
    datas=[
        ('NIST_BE.parquet', '.'),  # Include NIST_BE.parquet in bundle root
    ],
    hiddenimports=[
        'tkinter',
        'pandas',
        'numpy',
        'matplotlib',
        'pyperclip',
        'platform'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='KherveDB',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

app = BUNDLE(
    exe,
    name='KherveDB.app',
    icon=None,
    bundle_identifier='com.kherve.khervedb',
    info_plist={
        'NSPrincipalClass': 'NSApplication',
        'NSAppleScriptEnabled': False,
        'CFBundleDocumentTypes': [],
        'NSRequiresAquaSystemAppearance': 'No'
    },
)
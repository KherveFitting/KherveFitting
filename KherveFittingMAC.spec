# -*- mode: python ; coding: utf-8 -*-

import os
from PyInstaller.utils.hooks import collect_all

block_cipher = None

# Collect wx components
wxwidgets = collect_all('wx')

# Gather all .py files from the libraries folder
library_files = [(os.path.join('libraries', f), 'libraries') for f in os.listdir('libraries') if f.endswith('.py')]

a = Analysis(['KherveFitting.py'],
             pathex=['.'],
             binaries=wxwidgets[1],  # Add wx binaries
             datas=[
                 ('KherveFitting_library.json', '.'),
                 ('KherveFitting_library.xlsx', '.'),
                 ('NIST_BE.parquet', '.'),
                 ('Manual.pdf', '.'),
                 (os.path.join('libraries', 'Images'), os.path.join('libraries', 'Images')),
                 (os.path.join('libraries', 'Icons'), os.path.join('libraries', 'Icons')),
                 ('Icons', 'Icons'),
             ] + library_files + wxwidgets[0],  # Add wx datas
             hiddenimports=['wx', 'numpy', 'matplotlib', 'pandas', 'openpyxl', 'lmfit', 'scipy',
                           'docx', 'vamas', 'yadg.extractors.phi.spe', 'psutil'] + wxwidgets[2],  # Add wx hiddenimports
             hookspath=[],
             hooksconfig={},
             runtime_hooks=[],
             excludes=['PyQt5', 'PySide2', 'customtkinter'],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# For Mac, we need to use COLLECT and BUNDLE
exe = EXE(pyz,
          a.scripts,
          [],  # No binaries here for Mac bundle
          exclude_binaries=True,  # Exclude binaries for Mac
          name='KherveFitting',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False,
          argv_emulation=True,  # Important for Mac
          icon='Icons/Icon.ico')  # Will work but .icns would be better

# COLLECT gathers all the pieces
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='KherveFitting')

# BUNDLE creates a Mac app bundle
app = BUNDLE(coll,
             name='KherveFitting.app',
             icon='Icons/Icon.ico',  # Can use .ico, but .icns is preferred
             bundle_identifier='com.khervefitting.app',
             info_plist={
                 'CFBundleName': 'KherveFitting',
                 'CFBundleDisplayName': 'KherveFitting',
                 'CFBundleVersion': '1.4',
                 'CFBundleShortVersionString': '1.4',
                 'NSHumanReadableCopyright': '© 2025',
                 'CFBundleDocumentTypes': [
                     {
                         'CFBundleTypeExtensions': ['xlsx', 'vms', 'kal', 'avg', 'spe'],
                         'CFBundleTypeName': 'XPS Data File',
                         'CFBundleTypeRole': 'Editor',
                     }
                 ]
             })
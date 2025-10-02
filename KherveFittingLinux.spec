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
             binaries=wxwidgets[1],
             datas=[
                 ('LICENSE', '.'),
                 ('THIRD_PARTY_LICENSES.txt', '.'),
                 ('KherveFitting_library.json', '.'),
                 (os.path.join('libraries', 'Images'), os.path.join('libraries', 'Images')),
                 (os.path.join('libraries', 'Icons'), os.path.join('libraries', 'Icons')),
                 (os.path.join('libraries', 'ToolsMenu', 'Icons'), os.path.join('libraries', 'ToolsMenu', 'Icons')),
                 (os.path.join('libraries', 'ViewMenu', 'Icons'), os.path.join('libraries', 'ViewMenu', 'Icons')),
                 ('Icons', 'Icons'),
             ] + library_files + wxwidgets[0],
             hiddenimports=['wx', 'numpy', 'matplotlib', 'pandas', 'openpyxl', 'lmfit', 'scipy','matplotlib.backends._backend_agg', 'matplotlib.backends.backend_wxagg', 'matplotlib.backends.backend_svg'] + wxwidgets[2],
             hookspath=[],
             hooksconfig={},
             runtime_hooks=[],
             excludes=['PyQt5', 'PySide2', 'customtkinter'],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          [],
          exclude_binaries=True,
          name='KherveFitting',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False,
          disable_windowed_traceback=False,
          target_arch=None,
          codesign_identity=None,
          entitlements_file=None)

coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='KherveFitting')
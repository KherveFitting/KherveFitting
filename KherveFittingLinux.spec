# -*- mode: python ; coding: utf-8 -*-

import os
from PyInstaller.utils.hooks import collect_all

block_cipher = None

# Collect wx components
wxwidgets = collect_all('wx')

# Gather all .py files from the libraries folder
library_files = [(os.path.join('libraries', f), 'libraries') for f in os.listdir('libraries') if f.endswith('.py')]

# config.json is user generated and not in the repository, ship it only if present
optional_files = [(f, '.') for f in ['config.json'] if os.path.exists(f)]

a = Analysis(['KherveFitting.py'],
             pathex=['.'],
             binaries=wxwidgets[1],
             datas=[
                 ('LICENSE', '.'),
                 ('THIRD_PARTY_LICENSES.txt', '.'),
                 ('KherveFitting_library.json', '.'),
                 ('KherveFitting_library.xlsx', '.'),
                 ('KherveFitting_library.parquet', '.'),
                 ('NIST_BE.parquet', '.'),
                 ('Manual.pdf', '.'),
                 (os.path.join('libraries', 'Images'), os.path.join('libraries', 'Images')),
                 (os.path.join('libraries', 'Icons'), os.path.join('libraries', 'Icons')),
                 (os.path.join('libraries', 'ToolsMenu', 'Icons'), os.path.join('libraries', 'ToolsMenu', 'Icons')),
                 (os.path.join('libraries', 'ViewMenu', 'Icons'), os.path.join('libraries', 'ViewMenu', 'Icons')),
                 (os.path.join('libraries', 'HelpMenu', 'Images'), os.path.join('libraries', 'HelpMenu', 'Images')),
                 (os.path.join('libraries', 'Games', 'Cards'), os.path.join('libraries', 'Games', 'Cards')),
                 ('Icons', 'Icons'),
             ] + optional_files + library_files + wxwidgets[0],
             hiddenimports=['wx', 'numpy', 'matplotlib', 'pandas', 'openpyxl', 'lmfit', 'scipy',
                            'matplotlib.backends._backend_agg', 'matplotlib.backends.backend_wxagg',
                            'matplotlib.backends.backend_svg',
                            'docx', 'vamas', 'yadg.extractors.phi.spe', 'psutil'] + wxwidgets[2],
             hookspath=[],
             hooksconfig={},
             runtime_hooks=[],
             excludes=['PyQt5', 'PySide2', 'customtkinter'],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)

# PyInstaller collects the GTK stack that wxWidgets links against, but not the
# gdk-pixbuf loaders and icon themes that belong with it, so the bundled copy
# cannot draw some assets. The Debian package depends on the system GTK, so drop
# those libraries and let the application use the desktop's own theme. Remove
# this block to get a fully self contained bundle like the Windows build.
SYSTEM_GTK_PREFIXES = (
    'libgtk-3', 'libgdk-3', 'libgdk_pixbuf-2.0', 'libgio-2.0', 'libglib-2.0',
    'libgobject-2.0', 'libgmodule-2.0', 'libpango', 'libcairo', 'libatk',
    'libepoxy', 'libharfbuzz', 'libfontconfig', 'libfreetype',
)
a.binaries = [b for b in a.binaries
              if not os.path.basename(b[0]).startswith(SYSTEM_GTK_PREFIXES)]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          [],
          exclude_binaries=True,
          name='KherveFitting',
          # Keep the bundled data next to the executable rather than in _internal:
          # the application looks for its resources relative to sys.executable on
          # every platform except macOS.
          contents_directory='.',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=False,
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
               upx=False,
               upx_exclude=[],
               name='KherveFitting')

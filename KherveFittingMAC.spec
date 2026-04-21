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
                 ('KherveFitting_library.parquet', '.'),
                 ('config.json', '.'),
                 ('NIST_BE.parquet', '.'),
                 ('Manual.pdf', '.'),
                 (os.path.join('libraries', 'Images'), os.path.join('libraries', 'Images')),
                 (os.path.join('libraries', 'Icons'), os.path.join('libraries', 'Icons')),
                 (os.path.join('libraries', 'ViewMenu','Icons'), os.path.join('libraries', 'ViewMenu','Icons')),
                 (os.path.join('libraries', 'HelpMenu','Images'), os.path.join('libraries', 'ViewMenu','Images')),
                 ('Icons', 'Icons'),
                 ('Data-Examples', 'Data-Examples'),
                 ('Backup', 'Backup'),
             ] + library_files + wxwidgets[0],
             hiddenimports=['wx', 'numpy', 'matplotlib', 'pandas', 'openpyxl', 'lmfit', 'scipy',
                          'docx', 'vamas', 'yadg.extractors.phi.spe', 'psutil',
                          'sklearn', 'sklearn.utils._cython_blas', 'sklearn.neighbors.typedefs',
                          'sklearn.neighbors.quad_tree', 'sklearn.tree._utils', 'sklearn.utils._weight_vector',
                          'sklearn.externals', 'sklearn.externals.array_api_compat',
                          'sklearn.externals.array_api_compat.numpy', 'sklearn.externals.array_api_compat.numpy.fft',
                          'sklearn.decomposition', 'sklearn.preprocessing',
                          'sklearn.__check_build', 'sklearn._cyutility', 'sklearn.utils._isfinite'] + wxwidgets[2],
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
          [],
          exclude_binaries=True,
          name='KherveFitting',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False,
          disable_windowed_traceback=False,
          argv_emulation=True,
          target_arch=None,
          codesign_identity=None,
          entitlements_file=None,
          icon='icons/Icon.icns')

# COLLECT gathers all the pieces
coll = COLLECT(exe,
               a.binaries,
               #a.zipfiles,
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
                 'CFBundleVersion': '1.7',
                 'CFBundleShortVersionString': '1.7',
                 'NSHumanReadableCopyright': '© 2025',
                 'CFBundleDocumentTypes': [
                     {
                         'CFBundleTypeExtensions': ['xlsx', 'vms', 'kal', 'avg', 'spe'],
                         'CFBundleTypeName': 'XPS Data File',
                         'CFBundleTypeRole': 'Editor',
                     }
                 ]
             })
# -*- mode: python ; coding: utf-8 -*-

import os
from kivy_deps import sdl2, glew


block_cipher = None


a = Analysis(['controlegeral.py'],
             pathex=[],
             binaries=[],
             datas=[('ControleGeral.kv','.'),('*.py','.'),('*.png','.'),('database_tolda.db','.'),('logoCN.ico','.')],
             hiddenimports=['win32timezone'],
             hookspath=[],
             hooksconfig={},
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)

pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)

exe = EXE(pyz,
          a.scripts, 
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          exclude_binaries=True,
          name='Controle Geral CN',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False,
          disable_windowed_traceback=False,
          target_arch=None,
          codesign_identity=None,
          entitlements_file=None,
          icon= 'logoCN.ico')

coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='Controle Geral CN')
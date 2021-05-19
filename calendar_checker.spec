# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['./calendar_checker/checker.py'],
             pathex=['./calendar_checker'],
             binaries=[],
             datas=[("./calendar_checker/data", "./data")],
             hiddenimports=['win32timezone'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          [],
          exclude_binaries=True,
          name='checker',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False ,
          icon=('calendar_checker/data/calendar.ico'))

coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='checker')

# -*- mode: python ; coding: utf-8 -*-


block_cipher = None

added_files = [
    ( 'theme\\forest-dark.tcl', 'theme' ),
    ('theme\\forest-dark\\*.png', 'theme\\forest-dark' ),
    ('ico\\s1_manager.*', 'ico')
]

a = Analysis(['s1_manager.py'],
             pathex=[],
             binaries=[],
             datas=added_files,
             hiddenimports=[],
             hookspath=[],
             hooksconfig={},
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False
             )
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,  
          [],
          name='s1_manager',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False,
          icon='ico\\s1_manager.ico',
          disable_windowed_traceback=False,
          target_arch=None,
          codesign_identity=None,
          entitlements_file=None,
          version='file_version_info.txt')

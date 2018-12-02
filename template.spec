# -*- mode: python -*-

block_cipher = None


a = Analysis(['main.py'],
             pathex=['Path/To/FOLDER/which/includes/main.py'],
             binaries=[],
             datas=[('Path/To/Python/Python37/Lib/site-packages/plotly/', './plotly/'),
                    ('Path/To/Python/Python37/Lib/site-packages/psutil/', './psutil/'),
                    ('Path/To/orca/', './orca/')],
             hiddenimports=['plotly', 'orca', 'psutil'],
             hookspath=[],
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
          [],
          exclude_binaries=True,
          name='main',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=True )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               name='main')

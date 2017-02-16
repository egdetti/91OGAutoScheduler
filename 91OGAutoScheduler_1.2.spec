# -*- mode: python -*-

block_cipher = None


a = Analysis(['91OGAutoScheduler_1.2.py'],
             pathex=['C:\\Users\\micha\\PycharmProjects\\91OGAutoScheduler'],
             binaries=[],
             datas=[('91OG.ico', '.')],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='91OGAutoScheduler_1.2',
          debug=False,
          strip=False,
          upx=True,
          console=False , icon='91OG.ico')

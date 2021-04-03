# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['main.py'],
             pathex=['.'],
             binaries=[],
             datas=[("st\icons","st\icons")],
             hiddenimports=["openpyxl",
                            "st.widgets.Viewer",
                            "st.widgets.CodeEditor",
                            "st.widgets.ExcelReader",
                            "st.widgets.ExcelWriter",
                            "st.widgets.Merger",
                            "st.widgets.Splitter"],
             hookspath=[],
             runtime_hooks=[],
             excludes=["sqlite3"],
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
          name='st',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False , 
          )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='st')

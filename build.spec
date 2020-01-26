a = Analysis(['hfest.py'],
             hiddenimports=['pony.orm.dbproviders.sqlite'])
pyz = PYZ(a.pure)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='hfest',
          upx=True)

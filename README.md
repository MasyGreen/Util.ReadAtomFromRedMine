#Convert RedMine Atom RSS to HTML and DOCX
1. Set settings in **config.cfg** (run once *.exe to create struct file)
   1. [host], IP or DNS RedMine name (example: **http://192.168.1.1**)
   2. [apikey], *RedMine - User - API key*, RESTAPI must by On (example: **aldjfoeiwgj9348gn348**)
   3. Export report from RedMine, format Atom to file *.css
   4. Run

Util crete file same name *.docx
* Delete empty string
* Delete some bug symbol
* LoweDown <H> (h1=h2, h2=h3 etc)

## Sample config.cfg
```
[Settings]
host = http://192.168.1.1
apikey = dq3inqgnqe8igqngninkkvekmviewrgir9384
```

# How use pyinstaller whit HtmlToDocx

1. Run in CMD

```
pyinstaller -F -i "Icon.ico" RedMineIssueToDocx.py
```

2. Edit *RedMineIssueToDocx.spec*
   1. Add at first line *import sys* and *from os import path*
   2. Update section in Analysis, where *c:\\Python38\\Lib\\site-packages* is your path to *docx/templates* folder:

```
datas=[(path.join("c:\\Python38\\Lib\\site-packages","docx","templates"), "docx/templates")]
```
   5. Run in CMD

```
pyinstaller ReadAtomFromRedMine.spec
```
   6. Result in **dist\ReadAtomFromRedMine.exe**

## Sample ReadAtomFromRedMine.spec
```
# -*- mode: python ; coding: utf-8 -*-

import sys
from os import path

block_cipher = None


a = Analysis(['ReadAtomFromRedMine.py'],
             pathex=[],
             binaries=[],
             datas=[(path.join("c:\\Python38\\Lib\\site-packages","docx","templates"), "docx/templates")],
             hiddenimports=[],
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
          name='ReadAtomFromRedMine',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=True,
          disable_windowed_traceback=False,
          target_arch=None,
          codesign_identity=None,
          entitlements_file=None , icon='Icon.ico')
```
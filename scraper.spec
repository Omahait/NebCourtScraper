# -*- mode: python ; coding: utf-8 -*-

import os
from PyInstaller.utils.hooks import collect_data_files

APP_NAME = "Nebraska Courts E-Services Scraper V2"
SCRIPT   = "scraper.py"

# Hidden imports that tkinter/tkcalendar/lxml/babel need at runtime
hidden = [
    "babel.localedata",
    "tkcalendar",
    "lxml.etree",
    "lxml._elementpath",
]

# 1) Pull in all Babel data (for tkcalendar locales)
datas = collect_data_files("babel", include_py_files=False)

# 2) Add the entire assets folder recursively (icons + themes)
assets_root = "assets"
if os.path.isdir(assets_root):
    for root, _, files in os.walk(assets_root):
        for f in files:
            src = os.path.join(root, f)               # source on disk
            dst = os.path.join(root)                   # destination folder inside bundle
            datas.append((src, dst))

a = Analysis(
    [SCRIPT],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    name=APP_NAME,
    icon=os.path.join("assets", "app.ico"),  # EXE icon (Explorer/Properties)
    console=False,                            # windowed app
    disable_windowed_traceback=False,
    target_arch=None,
    version=None,
    uac_admin=False,
    uac_uiaccess=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name=APP_NAME,
)

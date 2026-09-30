# -*- mode: python ; coding: utf-8 -*-
# Build:  python -m PyInstaller training_planner_gui.spec
from PyInstaller.utils.hooks import collect_all, collect_data_files

datas = collect_data_files('sv_ttk') + collect_data_files('tkcalendar')
binaries = []
hiddenimports = ['babel.numbers', 'tkcalendar', 'sv_ttk', 'darkdetect']
for pkg in ('numpy', 'pandas'):
    d, b, h = collect_all(pkg)
    datas += d; binaries += b; hiddenimports += h

a = Analysis(
    ['garmin_planner.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz, a.scripts, a.binaries, a.datas, [],
    name='GarminTrainingPlanner',
    debug=False,
    strip=False,
    upx=True,
    console=False,
    argv_emulation=False,
)
app = BUNDLE(exe, name='GarminTrainingPlanner.app', bundle_identifier=None)  # solo macOS

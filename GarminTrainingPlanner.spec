# -*- mode: python ; coding: utf-8 -*-
# Crea un unico eseguibile con tutte le dipendenze (Python compreso).
#   Windows:  build_exe.bat          ->  dist\GarminTrainingPlanner.exe
#   macOS/Linux:  ./build_app.sh     ->  dist/GarminTrainingPlanner(.app)
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

datas = [("assets/app.png", "assets"), ("assets/app.ico", "assets")]
datas += collect_data_files("sv_ttk")        # file del tema grafico
datas += collect_data_files("tkcalendar")
datas += collect_data_files("babel")         # nomi di giorni/mesi in italiano per il calendario

hiddenimports = ["babel.numbers", "babel.dates", "tkcalendar", "sv_ttk", "darkdetect"]
hiddenimports += collect_submodules("garth")  # client Garmin Connect

# moduli pesanti non usati dall'app (eseguibile piu' piccolo)
excludes = ["matplotlib", "scipy", "IPython", "notebook", "pytest", "PyQt5", "PyQt6", "logfire",
            "PySide2", "PySide6", "sphinx", "docutils"]

a = Analysis(
    ["garmin_planner.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=["build_hooks/rthook_env.py"],
    excludes=excludes,
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="GarminTrainingPlanner",
    icon="assets/app.ico",
    debug=False,
    strip=False,
    upx=False,          # UPX fa scattare piu' spesso i falsi allarmi degli antivirus
    console=False,      # nessuna finestra nera; gli errori finiscono in ~/.garminplanner/errore_avvio.log
    runtime_tmpdir=None,
)

if sys.platform == "darwin":
    app = BUNDLE(exe, name="GarminTrainingPlanner.app", icon="assets/app.png",
                 bundle_identifier="it.garminplanner.app")

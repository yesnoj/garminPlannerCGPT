#!/usr/bin/env bash
# Crea l'eseguibile su macOS (dist/GarminTrainingPlanner.app) o Linux (dist/GarminTrainingPlanner)
set -euo pipefail
cd "$(dirname "$0")"
PY=${PYTHON:-python3}
[ -d .venv-build ] || "$PY" -m venv .venv-build
source .venv-build/bin/activate
python -m pip install --upgrade pip >/dev/null
python -m pip install -r requirements.txt pyinstaller
python -m PyInstaller --noconfirm --clean GarminTrainingPlanner.spec
echo "FATTO: dist/"
ls -la dist

# Mapala

Outil simple de mapping template ↔ source avec concaténation de colonnes.

## Installation

```bash
python -m venv .venv
source .venv/bin/activate
pip install -e .
```

## Lancer l'app

```bash
mapala
```

## Build

Windows (.exe):
```bash
pip install -e .[build,formats]
python build_exe.py
```

macOS (.app / .dmg):
```bash
pip install -e .[build,formats]
python build_macos_app.py
./build_macos_dmg.sh
```

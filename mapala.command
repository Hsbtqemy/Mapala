#!/bin/bash
# Lance Mapala (macOS)

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
APP_DIR="$SCRIPT_DIR"
VENV_DIR="$APP_DIR/.venv"
VENV_PY="$VENV_DIR/bin/python"

cd "$APP_DIR"

if [ ! -x "$VENV_PY" ]; then
  echo "Création de l'environnement Python..."
  if command -v python3 >/dev/null 2>&1; then
    python3 -m venv "$VENV_DIR"
  else
    echo "python3 introuvable. Installez Python 3.11+ et relancez."
    exit 1
  fi
  echo "Installation des dépendances..."
  "$VENV_PY" -m pip install --upgrade pip
  "$VENV_PY" -m pip install -e .
fi

exec "$VENV_PY" -m mapala

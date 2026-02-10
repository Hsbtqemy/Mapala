#!/bin/bash
# Lance Mapala (macOS)

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
APP_DIR="$SCRIPT_DIR"

if [ -f "$APP_DIR/.venv/bin/activate" ]; then
  source "$APP_DIR/.venv/bin/activate"
fi

python -m mapala

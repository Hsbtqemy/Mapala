#!/usr/bin/env python
"""
Script de build pour créer les exécutables Windows (Mapala GUI).

Usage:
    pip install -e ".[build,formats]"
    python build_exe.py

Sorties:
    - dist/Mapala/ (version portable + DLLs)
    - dist/MapalaStandalone.exe (version autonome)
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path


def _run(cmd: list[str], root: Path) -> int:
    print("Execution:", " ".join(cmd))
    result = subprocess.run(cmd, cwd=root)
    return result.returncode


def main() -> int:
    root = Path(__file__).parent.resolve()
    src = root / "src"
    app = root / "src" / "mapala" / "app.py"

    base_cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--windowed",
        "--clean",
        "--noconfirm",
        f"--paths={src}",
        str(app),
        "--exclude-module=tkinter",
        "--exclude-module=matplotlib",
        "--exclude-module=scipy",
        "--exclude-module=numpy.testing",
    ]

    # Portable (onedir)
    cmd_portable = [*base_cmd, "--name=Mapala", "--onedir"]
    if _run(cmd_portable, root) != 0:
        return 1
    print("\nBuild reussi. Portable dans: dist/Mapala/Mapala.exe")

    # Standalone (onefile)
    cmd_standalone = [*base_cmd, "--name=MapalaStandalone", "--onefile"]
    if _run(cmd_standalone, root) != 0:
        return 1
    print("\nBuild reussi. Standalone dans: dist/MapalaStandalone.exe")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

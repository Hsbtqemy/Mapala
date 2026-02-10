#!/usr/bin/env python
"""
Script de build pour créer un exécutable Windows autonome (Mapala GUI).

Usage:
    pip install -e ".[build,formats]"
    python build_exe.py

Le .exe et les DLL seront dans dist/Mapala/
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path


def main() -> int:
    root = Path(__file__).parent.resolve()
    src = root / "src"

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--name=Mapala",
        "--windowed",
        "--onedir",
        "--clean",
        "--noconfirm",
        f"--paths={src}",
        str(root / "src" / "mapala" / "app.py"),
        "--exclude-module=tkinter",
        "--exclude-module=matplotlib",
        "--exclude-module=scipy",
        "--exclude-module=numpy.testing",
    ]

    print("Exécution:", " ".join(cmd))
    result = subprocess.run(cmd, cwd=root)
    if result.returncode == 0:
        print("\nBuild reussi. Executable dans: dist/Mapala/Mapala.exe")
    return result.returncode


if __name__ == "__main__":
    raise SystemExit(main())

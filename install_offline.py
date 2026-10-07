"""Install project packages from the local packages/ folder. No network is used.

    python install_offline.py
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
REQUIREMENTS = ROOT / "requirements.txt"
DEST = ROOT / "packages"


def main() -> None:
    if sys.version_info[:2] != (3, 14):
        raise SystemExit(
            f"Install Python 3.14.7 (64-bit). Current interpreter is {sys.version.split()[0]}."
        )
    if not DEST.is_dir() or not any(DEST.glob("*.whl")):
        raise SystemExit(
            f"No wheels in {DEST}. On a connected machine run: python download_packages.py"
        )
    command = [
        sys.executable,
        "-m",
        "pip",
        "install",
        "--no-index",
        "--find-links",
        str(DEST),
        "-r",
        str(REQUIREMENTS),
    ]
    print("Installing from", DEST)
    subprocess.check_call(command)
    print("Offline install finished.")


if __name__ == "__main__":
    main()

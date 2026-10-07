"""Download installers for this project so a later machine can install with no network.

Run once on a computer that has internet and the same Python as the offline machine:

    python download_packages.py

Wheels land in the packages/ folder. Copy that folder with the project.
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
            f"This bundle is for Python 3.14. Current interpreter is {sys.version.split()[0]}."
        )
    DEST.mkdir(exist_ok=True)
    command = [
        sys.executable,
        "-m",
        "pip",
        "download",
        "-r",
        str(REQUIREMENTS),
        "-d",
        str(DEST),
    ]
    print("Downloading wheels into", DEST)
    subprocess.check_call(command)
    stamp = DEST / "PYTHON_VERSION.txt"
    stamp.write_text(
        f"{sys.version}\nplatform={sys.platform}\n",
        encoding="utf-8",
    )
    count = len(list(DEST.glob("*.whl"))) + len(list(DEST.glob("*.tar.gz")))
    print(f"Saved {count} package files to {DEST}")


if __name__ == "__main__":
    main()

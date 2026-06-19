from __future__ import annotations

import subprocess
import sys
from pathlib import Path


def main() -> int:
    app_path = Path(__file__).resolve().parent / "front.py"

    if not app_path.exists():
        print(f"front.py wurde nicht gefunden: {app_path}")
        return 1

    command = [
        sys.executable,
        "-m",
        "streamlit",
        "run",
        str(app_path),
    ]

    return subprocess.call(command)


if __name__ == "__main__":
    raise SystemExit(main())
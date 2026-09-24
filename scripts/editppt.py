"""Run this skill's bundled editppt code, ignoring older global installations."""

import os
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "cli"))
os.environ["IMAGE_TO_EDITABLE_PPT_CLI_PROG"] = f'"{sys.executable}" "{Path(__file__).resolve()}"'

from editppt.cli import main


if __name__ == "__main__":
    main()

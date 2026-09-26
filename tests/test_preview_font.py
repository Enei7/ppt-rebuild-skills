"""Windows previews should use a scalable font, not Pillow's tiny bitmap fallback."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "cli" / "editppt" / "runtime"))

from build_pptx_from_manifest import choose_preview_font


def main():
    if sys.platform == "win32":
        from PIL import ImageFont

        font = choose_preview_font(None)
        assert font and Path(font).is_file()
        assert ImageFont.truetype(font, 48).getbbox("Topic")[3] > 30


if __name__ == "__main__":
    main()

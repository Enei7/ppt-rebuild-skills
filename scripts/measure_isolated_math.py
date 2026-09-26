"""Measure complete isolated formula ink, without clipping to its source box."""

import argparse
import json
from pathlib import Path

import numpy as np
from PIL import Image

from measure_math_geometry import ink_box


def compare(source, rendered, box):
    if source.shape != rendered.shape:
        raise ValueError("Source and isolated render dimensions differ")
    x, y, width, height = map(int, box)
    if min(x, y) < 0 or min(width, height) <= 0 or x + width > source.shape[1] or y + height > source.shape[0]:
        raise ValueError(f"Invalid formula source box: {box}")
    original = ink_box(source[y:y + height, x:x + width])
    current = ink_box(rendered)
    if original is None or current is None:
        raise ValueError("Missing source or rendered formula ink")
    original = [original[0] + x, original[1] + y, original[2] + x, original[3] + y]
    if current[0] == 0 or current[1] == 0 or current[2] == rendered.shape[1] or current[3] == rendered.shape[0]:
        raise ValueError("Native ink touches canvas edge; resolve clipping before measurement")
    ratio = min((original[2] - original[0]) / (current[2] - current[0]),
                (original[3] - original[1]) / (current[3] - current[1]))
    return {"ratio": ratio, "dx_px": (original[0] + original[2] - current[0] - current[2]) / 2,
            "dy_px": (original[1] + original[3] - current[1] - current[3]) / 2,
            "source_ink": original, "rendered_ink": current,
            "source_size_px": [source.shape[1], source.shape[0]]}


def measure(run, folder):
    targets = json.loads((folder / "targets.json").read_text(encoding="utf-8-sig"))
    result = []
    for index, target in enumerate(targets):
        page = run / "pages" / f"page_{int(target['page']):03}"
        manifest = json.loads((page / "manifest.json").read_text(encoding="utf-8-sig"))
        matches = [f for f in manifest.get("formula_inventory", []) if f.get("id") == target["id"]]
        if len(matches) != 1:
            raise ValueError(f"Formula inventory id is not unique: {target}")
        source = np.asarray(Image.open(page / "source.png").convert("RGB"))
        rendered = np.asarray(Image.open(folder / f"formula_{index:04}.png").convert("RGB"))
        result.append({**target, **compare(source, rendered, matches[0]["box_px"])})
    return result


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--run", type=Path, required=True)
    parser.add_argument("--renders", type=Path, required=True)
    args = parser.parse_args()
    print(json.dumps(measure(args.run, args.renders), ensure_ascii=False))

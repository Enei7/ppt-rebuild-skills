"""Measure complete isolated formula ink, without clipping to its source box."""

import argparse
import json
from pathlib import Path

import numpy as np
from PIL import Image

from formula_source_geometry import (
    resolve_formula_source_geometry,
    translate_source_ink_to_target,
)
from measure_math_geometry import ink_box


SOURCE_INK_FILTERS = {"achromatic"}


def validate_source_ink_filter(source_ink_filter):
    if source_ink_filter is None:
        return None
    if (
        not isinstance(source_ink_filter, str)
        or source_ink_filter not in SOURCE_INK_FILTERS
    ):
        allowed = ", ".join(sorted(SOURCE_INK_FILTERS))
        raise ValueError(
            f"Invalid source_ink_filter {source_ink_filter!r}; expected one of: {allowed}"
        )
    return source_ink_filter


def filter_source_ink(pixels, source_ink_filter=None):
    source_ink_filter = validate_source_ink_filter(source_ink_filter)
    if source_ink_filter is None:
        return pixels
    rgb = pixels[:, :, :3].astype(np.int16, copy=False)
    chromatic = np.max(rgb, axis=2) - np.min(rgb, axis=2) > 60
    filtered = pixels.copy()
    filtered[chromatic, :3] = 255
    return filtered


def compare(source, rendered, source_box, target_box=None, source_ink_filter=None):
    if source.shape != rendered.shape:
        raise ValueError("Source and isolated render dimensions differ")
    if target_box is None:
        target_box = source_box
    x, y, width, height = map(int, source_box)
    if min(x, y) < 0 or min(width, height) <= 0 or x + width > source.shape[1] or y + height > source.shape[0]:
        raise ValueError(f"Invalid formula source box: {source_box}")
    source_ink_filter = validate_source_ink_filter(source_ink_filter)
    source_crop = filter_source_ink(
        source[y:y + height, x:x + width], source_ink_filter
    )
    original_in_source_box = ink_box(source_crop)
    current = ink_box(rendered)
    if original_in_source_box is None:
        suffix = (
            f" after source_ink_filter={source_ink_filter!r}"
            if source_ink_filter is not None else ""
        )
        raise ValueError(f"Missing source formula ink{suffix}")
    if current is None:
        raise ValueError("Missing rendered formula ink")
    expected_in_target_box = translate_source_ink_to_target(
        original_in_source_box, source_box, target_box
    )
    tx, ty, _, _ = map(float, target_box)
    original = [expected_in_target_box[0] + tx, expected_in_target_box[1] + ty,
                expected_in_target_box[2] + tx, expected_in_target_box[3] + ty]
    if current[0] == 0 or current[1] == 0 or current[2] == rendered.shape[1] or current[3] == rendered.shape[0]:
        raise ValueError("Native ink touches canvas edge; resolve clipping before measurement")
    ratio = min((original[2] - original[0]) / (current[2] - current[0]),
                (original[3] - original[1]) / (current[3] - current[1]))
    return {"ratio": ratio, "dx_px": (original[0] + original[2] - current[0] - current[2]) / 2,
            "dy_px": (original[1] + original[3] - current[1] - current[3]) / 2,
            "source_ink": original, "source_ink_in_source_box": original_in_source_box,
            "source_box_px": list(map(float, source_box)),
            "target_box_px": list(map(float, target_box)), "rendered_ink": current,
            "source_size_px": [source.shape[1], source.shape[0]]}


def measure(run, folder):
    targets = json.loads((folder / "targets.json").read_text(encoding="utf-8-sig"))
    result = []
    for index, target in enumerate(targets):
        source_ink_filter = validate_source_ink_filter(
            target.get("source_ink_filter")
        )
        page = run / "pages" / f"page_{int(target['page']):03}"
        manifest = json.loads((page / "manifest.json").read_text(encoding="utf-8-sig"))
        matches = [f for f in manifest.get("formula_inventory", []) if f.get("id") == target["id"]]
        if len(matches) != 1:
            raise ValueError(f"Formula inventory id is not unique: {target}")
        source = np.asarray(Image.open(page / "source.png").convert("RGB"))
        rendered = np.asarray(Image.open(folder / f"formula_{index:04}.png").convert("RGB"))
        geometry = resolve_formula_source_geometry(
            manifest, matches[0], int(target["page"]), source.shape[1::-1]
        )
        result.append({
            **target,
            **compare(
                source,
                rendered,
                geometry.source_box_px,
                geometry.target_box_px,
                source_ink_filter=source_ink_filter,
            ),
            "source_box_kind": geometry.source_kind,
            "source_box_evidence": list(geometry.evidence),
        })
    return result


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--run", type=Path, required=True)
    parser.add_argument("--renders", type=Path, required=True)
    args = parser.parse_args()
    print(json.dumps(measure(args.run, args.renders), ensure_ascii=False))

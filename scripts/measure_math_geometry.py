"""Measure native equation ink against the matching source-PDF crop."""

import argparse
import json
from pathlib import Path

import numpy as np
from PIL import Image


def ink_box(pixels):
    dark = np.min(pixels[:, :, :3], axis=2) < 180
    ys, xs = np.nonzero(dark)
    if len(xs) < 8:
        return None
    return [int(xs.min()), int(ys.min()), int(xs.max()) + 1, int(ys.max()) + 1]


def native_formula(formula):
    return isinstance(formula, dict) and not str(formula.get("decision", "")).startswith(("embedded-", "source-embedded-"))


def measure(run, renders=None):
    rows = []
    for page_dir in sorted((run / "pages").glob("page_*")):
        manifest = json.loads((page_dir / "manifest.json").read_text(encoding="utf-8"))
        page = int(page_dir.name.split("_")[-1])
        source = np.asarray(Image.open(page_dir / "source.png").convert("RGB"))
        rendered = None if renders is None else np.asarray(Image.open(renders / f"page_{page:03}.png").convert("RGB"))
        if rendered is not None and source.shape != rendered.shape:
            raise ValueError(f"Page {page}: source/render size differs")
        for formula in manifest.get("formula_inventory", []):
            if not native_formula(formula):
                continue
            x, y, w, h = map(int, formula["box_px"])
            original = ink_box(source[y : y + h, x : x + w])
            current = None if rendered is None else ink_box(rendered[y : y + h, x : x + w])
            row = {"page": page, "id": formula["id"], "latex": formula.get("latex", ""), "box_px": [x, y, w, h],
                   "source_ink": original, "rendered_ink": current,
                   "source_size_px": list(source.shape[1::-1])}
            if original and current:
                source_width = original[2] - original[0]
                rendered_width = current[2] - current[0]
                row["width_ratio"] = round(source_width / rendered_width, 5)
                row["dx_px"] = round((original[0] + original[2] - current[0] - current[2]) / 2, 2)
                row["dy_px"] = round((original[1] + original[3] - current[1] - current[3]) / 2, 2)
                row["edge_touch"] = current[0] <= 1 or current[1] <= 1 or current[2] >= w - 1 or current[3] >= h - 1
            rows.append(row)
    return rows


def self_check():
    blank = np.full((30, 40, 3), 255, dtype=np.uint8)
    blank[7:18, 9:24] = 0
    assert ink_box(blank) == [9, 7, 24, 18]
    assert ink_box(np.full_like(blank, 255)) is None
    assert native_formula({"decision": "source-exact-formula-crop"})
    assert not native_formula({"decision": "embedded-editable-text-run"})


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--self-check", action="store_true")
    parser.add_argument("--run", type=Path)
    parser.add_argument("--renders", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    if args.self_check:
        self_check()
        print("self-check passed")
    else:
        if not all((args.run, args.output)):
            parser.error("--run and --output are required")
        rows = measure(args.run, args.renders)
        args.output.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"measured={len(rows)} missing_source_ink={sum(not r['source_ink'] for r in rows)} output={args.output}")

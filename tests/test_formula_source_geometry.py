import json
import sys
import tempfile
from pathlib import Path

import numpy as np
from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from formula_source_geometry import (
    FormulaSourceGeometryError,
    resolve_formula_source_geometry,
    translate_source_ink_to_target,
)
from measure_isolated_math import measure as measure_isolated
from measure_math_geometry import measure
from measure_isolated_math import compare
from review_formula_crops import export_review


def manifest(source_crop=None, source_page=1, visual_crop=None):
    formula = {
        "id": "formula_1",
        "box_px": [100, 50, 80, 40],
        "image": "assets/formula_1.png",
        "replace_image_id": "formula_1",
    }
    result = {
        "formula_inventory": [formula],
        "images": [{"id": "formula_1", "path": "assets/formula_1.png", "box_px": formula["box_px"]}],
        "asset_provenance": [],
        "visual_inventory": [],
    }
    if source_crop is not None:
        result["asset_provenance"].append({
            "path": "assets/formula_1.png",
            "source": "source.png",
            "source_type": "user-provided",
            "exact_source_crop": True,
            "source_page": source_page,
            "source_crop_px": source_crop,
        })
    if visual_crop is not None:
        result["visual_inventory"].append({
            "id": "formula_1",
            "path": "assets/formula_1.png",
            "exact_source_crop": True,
            "source_page": source_page,
            "source_crop_px": visual_crop,
        })
    return result


def expect_error(action, text):
    try:
        action()
    except FormulaSourceGeometryError as exc:
        assert text in str(exc), str(exc)
    else:
        raise AssertionError(f"Expected FormulaSourceGeometryError containing {text!r}")


def test_legacy_source_equals_target_regression():
    data = manifest()
    resolved = resolve_formula_source_geometry(data, data["formula_inventory"][0], 1, [300, 200])
    assert resolved.source_box_px == resolved.target_box_px == (100.0, 50.0, 80.0, 40.0)
    assert translate_source_ink_to_target([10, 5, 60, 30], resolved.source_box_px, resolved.target_box_px) == [10.0, 5.0, 60.0, 30.0]


def test_exact_crop_is_resolved_and_center_translated_without_scaling():
    data = manifest([20, 20, 100, 30], visual_crop=[20, 20, 100, 30])
    resolved = resolve_formula_source_geometry(data, data["formula_inventory"][0], 1, [300, 200])
    assert resolved.source_box_px == (20.0, 20.0, 100.0, 30.0)
    # target is 20 px narrower and 10 px taller: center translation is (-10,+5)
    assert translate_source_ink_to_target([10, 4, 90, 26], resolved.source_box_px, resolved.target_box_px) == [0.0, 9.0, 80.0, 31.0]


def test_relocated_blank_target_still_measures_real_source_ink():
    with tempfile.TemporaryDirectory() as temp_dir:
        run = Path(temp_dir)
        page = run / "pages" / "page_001"
        page.mkdir(parents=True)
        data = manifest([20, 20, 100, 30], visual_crop=[20, 20, 100, 30])
        (page / "manifest.json").write_text(json.dumps(data), encoding="utf-8")
        source = np.full((200, 300, 3), 255, dtype=np.uint8)
        source[26:42, 35:95] = 0
        # formula target [100,50,80,40] is intentionally blank in source.png.
        Image.fromarray(source).save(page / "source.png")
        renders = run / "renders"
        renders.mkdir()
        rendered = np.full_like(source, 255)
        # Expected ink is source ink translated center-to-center into the target.
        rendered[61:77, 105:165] = 0
        Image.fromarray(rendered).save(renders / "page_001.png")
        rows = measure(run, renders)
        assert len(rows) == 1
        row = rows[0]
        assert row["source_box_px"] == [20.0, 20.0, 100.0, 30.0]
        assert row["target_box_px"] == [100.0, 50.0, 80.0, 40.0]
        assert row["source_ink_in_source_box"] == [15, 6, 75, 22]
        assert row["source_ink"] == [5.0, 11.0, 65.0, 27.0]
        assert row["rendered_ink"] == [5, 11, 65, 27]
        assert row["width_ratio"] == 1.0
        assert row["dx_px"] == 0 and row["dy_px"] == 0
        isolated = compare(source, rendered, [20, 20, 100, 30], [100, 50, 80, 40])
        assert isolated['source_ink'] == [105.0, 61.0, 165.0, 77.0]
        assert isolated['ratio'] == 1.0
        assert isolated['dx_px'] == 0 and isolated['dy_px'] == 0


def test_mismatched_source_page_fails():
    data = manifest([20, 20, 100, 30], source_page=2)
    expect_error(
        lambda: resolve_formula_source_geometry(data, data["formula_inventory"][0], 1, [300, 200]),
        "does not match",
    )


def test_isolated_measure_uses_same_resolver_and_target_translation():
    with tempfile.TemporaryDirectory() as temp_dir:
        root = Path(temp_dir)
        run = root / "run"
        page = run / "pages" / "page_001"
        renders = root / "isolated"
        page.mkdir(parents=True)
        renders.mkdir()
        data = manifest([20, 20, 100, 30], visual_crop=[20, 20, 100, 30])
        (page / "manifest.json").write_text(json.dumps(data), encoding="utf-8")
        source = np.full((200, 300, 3), 255, dtype=np.uint8)
        source[26:42, 35:95] = 0
        Image.fromarray(source).save(page / "source.png")
        rendered = np.full_like(source, 255)
        rendered[61:77, 105:165] = 0
        Image.fromarray(rendered).save(renders / "formula_0000.png")
        (renders / "targets.json").write_text(
            json.dumps([{"page": 1, "id": "formula_1"}]), encoding="utf-8"
        )
        row = measure_isolated(run, renders)[0]
        assert row["source_box_px"] == [20.0, 20.0, 100.0, 30.0]
        assert row["target_box_px"] == [100.0, 50.0, 80.0, 40.0]
        assert row["ratio"] == 1.0 and row["dx_px"] == 0 and row["dy_px"] == 0


def test_crop_review_exports_actual_source_box_not_target_box():
    with tempfile.TemporaryDirectory() as temp_dir:
        root = Path(temp_dir)
        run = root / "run"
        page = run / "pages" / "page_001"
        output = root / "review"
        page.mkdir(parents=True)
        data = manifest([20, 20, 100, 30], visual_crop=[20, 20, 100, 30])
        (page / "manifest.json").write_text(json.dumps(data), encoding="utf-8")
        source = np.full((200, 300, 3), 255, dtype=np.uint8)
        source[26:42, 35:95] = 0
        Image.fromarray(source).save(page / "source.png")
        report = export_review(run, output)
        row = report["items"][0]
        assert row["box_px"] == [100.0, 50.0, 80.0, 40.0]
        assert row["source_box_px"] == [20.0, 20.0, 100.0, 30.0]
        assert Image.open(output / row["crop"]).size == (100, 30)


def test_conflicting_provenance_and_visual_crops_fail():
    data = manifest([20, 20, 100, 30], visual_crop=[21, 20, 100, 30])
    expect_error(
        lambda: resolve_formula_source_geometry(data, data["formula_inventory"][0], 1, [300, 200]),
        "ambiguous exact source crops",
    )


def main():
    test_legacy_source_equals_target_regression()
    test_exact_crop_is_resolved_and_center_translated_without_scaling()
    test_relocated_blank_target_still_measures_real_source_ink()
    test_mismatched_source_page_fails()
    test_isolated_measure_uses_same_resolver_and_target_translation()
    test_crop_review_exports_actual_source_box_not_target_box()
    test_conflicting_provenance_and_visual_crops_fail()


if __name__ == "__main__":
    main()

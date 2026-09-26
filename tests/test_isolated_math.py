import json
import sys
import tempfile
from pathlib import Path

import numpy as np
from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from measure_isolated_math import compare, measure


def expect_value_error(action, message):
    try:
        action()
    except ValueError as exc:
        assert message in str(exc), str(exc)
    else:
        raise AssertionError(f"Expected ValueError containing {message!r}")


def test_existing_measurement_and_guards():
    source = np.full((100, 150, 3), 255, dtype=np.uint8)
    rendered = source.copy()
    source[30:40, 50:70] = 0
    rendered[45:65, 80:120] = 0  # deliberately outside the source crop
    row = compare(source, rendered, [45, 25, 30, 20])
    assert row["ratio"] == 0.5
    assert row["dx_px"] == -40 and row["dy_px"] == -20
    for invalid in ([0, 0, 5, 5], [-1, 0, 30, 20], [100, 80, 100, 40]):
        expect_value_error(
            lambda invalid=invalid: compare(source, rendered, invalid),
            "source",
        )
    rendered[0:10, 0:10] = 0
    expect_value_error(
        lambda: compare(source, rendered, [45, 25, 30, 20]),
        "touches canvas edge",
    )


def annotated_source():
    source = np.full((100, 150, 3), 255, dtype=np.uint8)
    source[30:40, 50:70] = 0
    source[45:47, 40:100] = [255, 0, 0]
    source[50:52, 45:105] = [0, 160, 0]
    rendered = np.full_like(source, 255)
    rendered[45:65, 80:120] = 0
    return source, rendered


def test_achromatic_filter_excludes_colored_annotations():
    source, rendered = annotated_source()
    row = compare(
        source,
        rendered,
        [35, 25, 80, 35],
        source_ink_filter="achromatic",
    )
    assert row["source_ink"] == [50, 30, 70, 40]
    assert row["rendered_ink"] == [80, 45, 120, 65]
    assert row["ratio"] == 0.5
    assert row["dx_px"] == -40 and row["dy_px"] == -20


def test_default_keeps_colored_annotations():
    source, rendered = annotated_source()
    row = compare(source, rendered, [35, 25, 80, 35])
    assert row["source_ink"] == [40, 30, 105, 52]
    explicit_none = compare(
        source,
        rendered,
        [35, 25, 80, 35],
        source_ink_filter=None,
    )
    assert explicit_none == row


def test_achromatic_filter_retains_gray_formula_ink():
    source = np.full((80, 120, 3), 255, dtype=np.uint8)
    source[20:35, 30:55] = [120, 120, 120]
    source[40:42, 20:70] = [255, 0, 0]
    rendered = np.full_like(source, 255)
    rendered[25:40, 40:65] = 0
    row = compare(
        source,
        rendered,
        [15, 15, 60, 30],
        source_ink_filter="achromatic",
    )
    assert row["source_ink"] == [30, 20, 55, 35]


def test_achromatic_filter_rejects_all_colored_source():
    source = np.full((80, 120, 3), 255, dtype=np.uint8)
    source[20:35, 30:55] = [255, 0, 0]
    rendered = np.full_like(source, 255)
    rendered[25:40, 40:65] = 0
    expect_value_error(
        lambda: compare(
            source,
            rendered,
            [15, 15, 60, 30],
            source_ink_filter="achromatic",
        ),
        "Missing source formula ink after source_ink_filter='achromatic'",
    )


def test_invalid_filter_fails_in_compare_and_targets_json():
    source, rendered = annotated_source()
    expect_value_error(
        lambda: compare(
            source,
            rendered,
            [35, 25, 80, 35],
            source_ink_filter="colored-annotation",
        ),
        "Invalid source_ink_filter",
    )
    expect_value_error(
        lambda: compare(
            source,
            rendered,
            [35, 25, 80, 35],
            source_ink_filter=["achromatic"],
        ),
        "Invalid source_ink_filter",
    )

    with tempfile.TemporaryDirectory() as temp_dir:
        root = Path(temp_dir)
        run = root / "run"
        page = run / "pages" / "page_001"
        renders = root / "renders"
        page.mkdir(parents=True)
        renders.mkdir()
        Image.fromarray(source).save(page / "source.png")
        Image.fromarray(rendered).save(renders / "formula_0000.png")
        (page / "manifest.json").write_text(
            json.dumps({
                "formula_inventory": [
                    {"id": "formula_1", "box_px": [35, 25, 80, 35]}
                ]
            }),
            encoding="utf-8",
        )
        targets = [{
            "page": 1,
            "id": "formula_1",
            "source_ink_filter": "achromatic",
        }]
        (renders / "targets.json").write_text(
            json.dumps(targets), encoding="utf-8"
        )
        rows = measure(run, renders)
        assert rows[0]["source_ink_filter"] == "achromatic"
        assert rows[0]["source_ink"] == [50, 30, 70, 40]

        targets[0]["source_ink_filter"] = "invalid"
        (renders / "targets.json").write_text(
            json.dumps(targets), encoding="utf-8"
        )
        expect_value_error(
            lambda: measure(run, renders),
            "Invalid source_ink_filter",
        )


def main():
    test_existing_measurement_and_guards()
    test_achromatic_filter_excludes_colored_annotations()
    test_default_keeps_colored_annotations()
    test_achromatic_filter_retains_gray_formula_ink()
    test_achromatic_filter_rejects_all_colored_source()
    test_invalid_filter_fails_in_compare_and_targets_json()


if __name__ == "__main__":
    main()

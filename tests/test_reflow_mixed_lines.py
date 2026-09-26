import json
import math
import sys
import tempfile
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from preflight_reflow_mixed_lines import (
    RecipeError,
    validate_recipes,
    validate_width_measurements,
)


SLIDE_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
 <p:cSld><p:spTree>
  <p:sp><p:nvSpPr><p:cNvPr id="1" name="prose_a"/></p:nvSpPr><p:txBody><a:p><a:r><a:t>Let </a:t></a:r></a:p></p:txBody></p:sp>
  <mc:AlternateContent>
   <mc:Choice Requires="a14"><p:sp><p:nvSpPr><p:cNvPr id="2" name="formula_x"/></p:nvSpPr><p:txBody><a:p><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></a:p></p:txBody></p:sp></mc:Choice>
   <mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="2" name="formula_x"/></p:nvSpPr><p:txBody><a:p><a:r><a:t> </a:t></a:r></a:p></p:txBody></p:sp></mc:Fallback>
  </mc:AlternateContent>
  <p:sp><p:nvSpPr><p:cNvPr id="3" name="dup_first"/></p:nvSpPr><p:txBody><a:p><a:r><a:t>same</a:t></a:r></a:p></p:txBody></p:sp>
  <p:sp><p:nvSpPr><p:cNvPr id="4" name="dup_second"/></p:nvSpPr><p:txBody><a:p><a:r><a:t>same</a:t></a:r></a:p></p:txBody></p:sp>
 </p:spTree></p:cSld>
</p:sld>"""


def expect_error(action, text):
    try:
        action()
    except RecipeError as exc:
        assert text in str(exc), str(exc)
    else:
        raise AssertionError(f"Expected RecipeError containing {text!r}")


def fixture(root: Path):
    run = root / "run"
    page = run / "pages" / "page_001"
    page.mkdir(parents=True)
    (page / "manifest.json").write_text(
        json.dumps(
            {
                "source": {"width_px": 2400, "height_px": 1350},
                "content_box": {"left": 0, "top": 0, "width": 13.333, "height": 7.5},
                "text_boxes": [
                    {"id": "prose_a"},
                    {"id": "dup_first"},
                    {"id": "dup_second"},
                ],
                "formula_inventory": [{"id": "formula_x", "decision": "source-formula-crop"}],
            }
        ),
        encoding="utf-8",
    )
    deck = root / "deck.pptx"
    with zipfile.ZipFile(deck, "w") as archive:
        archive.writestr("ppt/slides/slide1.xml", SLIDE_XML)
    return run, deck


def valid_recipe():
    return [
        {
            "id": "mixed-1",
            "page": 1,
            "text_font_pt": 24,
            "center_y_px": 500,
            "start_x_px": 200,
            "right_limit_px": 1800,
            "items": [
                {"shape": "prose_a", "expected_text": "Let "},
                {"shape": "formula_x", "expected_text": "x", "gap_before_px": 12},
                {"text": "same", "match_index": 1, "exact_expected": "same", "gap_before_px": 10},
            ],
        }
    ]


def test_valid_recipe_and_slide_order_match_index(run, deck):
    plan = validate_recipes(valid_recipe(), run, deck)
    assert [item["shape"] for item in plan["lines"][0]["items"]] == [
        "prose_a",
        "formula_x",
        "dup_second",
    ]
    assert [item["kind"] for item in plan["lines"][0]["items"]] == [
        "prose",
        "math",
        "prose",
    ]


def test_prose_does_not_require_manifest_id_and_math_font_override_is_ignored(run, deck):
    recipe = valid_recipe()
    recipe[0]["items"][1]["font_pt"] = 30
    plan = validate_recipes(recipe, run, deck)
    math = plan["lines"][0]["items"][1]
    assert math["kind"] == "math" and math["ignored_math_font_pt"] is True
    assert plan["warnings"] and "ignored" in plan["warnings"][0]


def test_required_types_and_finite_values(run, deck):
    missing = valid_recipe()
    del missing[0]["center_y_px"]
    expect_error(lambda: validate_recipes(missing, run, deck), "missing required fields")
    nonfinite = valid_recipe()
    nonfinite[0]["start_x_px"] = math.nan
    expect_error(lambda: validate_recipes(nonfinite, run, deck), "finite number")
    bad_page = valid_recipe()
    bad_page[0]["page"] = True
    expect_error(lambda: validate_recipes(bad_page, run, deck), "positive integer")


def test_ambiguous_text_requires_match_index(run, deck):
    recipe = valid_recipe()
    recipe[0]["items"][-1] = {"text": "same"}
    expect_error(lambda: validate_recipes(recipe, run, deck), "ambiguous")


def test_repeated_page_shape_fails_after_selector_resolution(run, deck):
    recipe = valid_recipe()
    recipe[0]["items"].append({"text": "Let "})
    expect_error(lambda: validate_recipes(recipe, run, deck), "repeated in line recipes")


def test_exact_expected_is_case_and_whitespace_sensitive(run, deck):
    recipe = valid_recipe()
    recipe[0]["items"][0]["expected_text"] = "Let"
    expect_error(lambda: validate_recipes(recipe, run, deck), "exact expected text differs")


def test_width_overflow_fails_without_resize(run, deck):
    plan = validate_recipes(valid_recipe(), run, deck)
    passing = [{
        "id": "mixed-1",
        "items": [
            {"shape": "prose_a", "width_px": 300},
            {"shape": "formula_x", "width_px": 400},
            {"shape": "dup_second", "width_px": 300},
        ],
    }]
    result = validate_width_measurements(plan, passing)
    assert result["status"] == "pass"
    overflow = json.loads(json.dumps(passing))
    overflow[0]["items"][-1]["width_px"] = 900
    expect_error(
        lambda: validate_width_measurements(plan, overflow),
        "no resize allowed",
    )


def test_runtime_omits_page_switch_when_pages_are_not_supplied():
    script = (Path(__file__).resolve().parents[1] / "scripts" / "reflow_mixed_lines.ps1").read_text(
        encoding="utf-8"
    )
    assert "if ($Pages)" in script
    assert "foreach ($page in $Pages)" in script
    assert "foreach ($page in @($Pages))" not in script


def test_all_overflowing_lines_are_reported(run, deck):
    plan = validate_recipes(valid_recipe(), run, deck)
    second = json.loads(json.dumps(plan['lines'][0]))
    second['id'] = 'mixed-2'
    plan['lines'].append(second)
    measurements = [
        {'id': name, 'items': [
            {'shape': 'prose_a', 'width_px': 300},
            {'shape': 'formula_x', 'width_px': 400},
            {'shape': 'dup_second', 'width_px': 900}]}
        for name in ('mixed-1', 'mixed-2')]
    try:
        validate_width_measurements(plan, measurements)
    except RecipeError as exc:
        assert "line 'mixed-1'" in str(exc)
        assert "line 'mixed-2'" in str(exc)
    else:
        raise AssertionError('Every overflow must fail validation')


def main():
    with tempfile.TemporaryDirectory() as temp_dir:
        run, deck = fixture(Path(temp_dir))
        test_valid_recipe_and_slide_order_match_index(run, deck)
        test_prose_does_not_require_manifest_id_and_math_font_override_is_ignored(run, deck)
        test_required_types_and_finite_values(run, deck)
        test_ambiguous_text_requires_match_index(run, deck)
        test_repeated_page_shape_fails_after_selector_resolution(run, deck)
        test_exact_expected_is_case_and_whitespace_sensitive(run, deck)
        test_width_overflow_fails_without_resize(run, deck)
        test_all_overflowing_lines_are_reported(run, deck)
        test_runtime_omits_page_switch_when_pages_are_not_supplied()


if __name__ == "__main__":
    main()

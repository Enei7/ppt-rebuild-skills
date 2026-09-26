"""COM-free validation for explicit source-reviewed mixed-line recipes."""

from __future__ import annotations

import argparse
import json
import math
import posixpath
import re
import sys
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS = {"p": P_NS, "a": A_NS, "m": M_NS, "r": R_NS, "mc": MC_NS}

LINE_FIELDS = {
    "id",
    "page",
    "text_font_pt",
    "center_y_px",
    "start_x_px",
    "right_limit_px",
    "items",
}
ITEM_FIELDS = {
    "shape",
    "text",
    "match_index",
    "expected_text",
    "exact_expected",
    "font_pt",
    "gap_before_px",
    "center_offset_px",
}


class RecipeError(ValueError):
    pass


def _load_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
        raise RecipeError(f"cannot read JSON {path}: {exc}") from exc


def _write_new_json(path: Path, payload: Any) -> None:
    if path.exists():
        raise RecipeError(f"refusing to overwrite: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def _finite(value: Any, label: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise RecipeError(f"{label} must be a finite number")
    result = float(value)
    if not math.isfinite(result):
        raise RecipeError(f"{label} must be a finite number")
    return result


def _positive_int(value: Any, label: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < 1:
        raise RecipeError(f"{label} must be a positive integer")
    return value


def load_recipe_document(path: Path) -> list[dict[str, Any]]:
    payload = _load_json(path)
    if isinstance(payload, dict):
        unknown = set(payload) - {"schema_version", "lines"}
        if unknown:
            raise RecipeError(f"top-level unknown fields: {sorted(unknown)}")
        if payload.get("schema_version", 1) != 1:
            raise RecipeError("schema_version must be 1")
        payload = payload.get("lines")
    if not isinstance(payload, list) or not payload:
        raise RecipeError("recipe must be a non-empty array or an object with non-empty lines")
    if not all(isinstance(line, dict) for line in payload):
        raise RecipeError("every line recipe must be an object")
    return payload


def _shape_text(shape: ET.Element) -> str:
    paragraphs = shape.findall("./p:txBody/a:p", NS)
    if not paragraphs:
        return "".join(
            node.text or ""
            for node in shape.iter()
            if node.tag in {f"{{{A_NS}}}t", f"{{{M_NS}}}t"}
        )
    values = []
    for paragraph in paragraphs:
        values.append(
            "".join(
                node.text or ""
                for node in paragraph.iter()
                if node.tag in {f"{{{A_NS}}}t", f"{{{M_NS}}}t"}
            )
        )
    return "\r".join(values)


def _active_slide_shapes(root: ET.Element) -> list[ET.Element]:
    """Return PowerPoint-visible p:sp nodes, choosing mc:Choice over Fallback."""

    shapes: list[ET.Element] = []

    def walk(node: ET.Element) -> None:
        if node.tag == f"{{{MC_NS}}}AlternateContent":
            choice = node.find(f"{{{MC_NS}}}Choice")
            branch = choice if choice is not None else node.find(f"{{{MC_NS}}}Fallback")
            if branch is not None:
                for child in list(branch):
                    walk(child)
            return
        if node.tag == f"{{{P_NS}}}sp":
            shapes.append(node)
            return
        for child in list(node):
            walk(child)

    walk(root)
    return shapes


def read_deck_shapes(deck: Path) -> tuple[int, dict[int, list[dict[str, Any]]]]:
    try:
        archive = zipfile.ZipFile(deck)
    except (OSError, zipfile.BadZipFile) as exc:
        raise RecipeError(f"cannot open PPTX {deck}: {exc}") from exc
    with archive:
        fallback_slide_files = sorted(
            (
                name
                for name in archive.namelist()
                if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)
            ),
            key=lambda name: int(re.search(r"(\d+)\.xml$", name).group(1)),
        )
        slide_files = fallback_slide_files
        presentation_name = "ppt/presentation.xml"
        rels_name = "ppt/_rels/presentation.xml.rels"
        if presentation_name in archive.namelist() and rels_name in archive.namelist():
            presentation = ET.fromstring(archive.read(presentation_name))
            relationships = ET.fromstring(archive.read(rels_name))
            targets = {
                rel.get("Id"): rel.get("Target")
                for rel in relationships.findall(f"{{{PKG_REL_NS}}}Relationship")
                if rel.get("Id") and rel.get("Target")
            }
            ordered = []
            for slide_id in presentation.findall("./p:sldIdLst/p:sldId", NS):
                rel_id = slide_id.get(f"{{{R_NS}}}id")
                target = targets.get(rel_id)
                if not target:
                    raise RecipeError(f"presentation slide relationship {rel_id!r} is missing")
                normalized = posixpath.normpath(posixpath.join("ppt", target))
                if normalized not in archive.namelist():
                    raise RecipeError(f"presentation slide target is missing: {normalized}")
                ordered.append(normalized)
            if ordered:
                slide_files = ordered
        pages: dict[int, list[dict[str, Any]]] = {}
        for page, filename in enumerate(slide_files, start=1):
            root = ET.fromstring(archive.read(filename))
            rows: list[dict[str, Any]] = []
            for order, shape in enumerate(_active_slide_shapes(root)):
                c_nv_pr = shape.find("./p:nvSpPr/p:cNvPr", NS)
                if c_nv_pr is None or not c_nv_pr.get("name"):
                    continue
                name = c_nv_pr.get("name")
                rows.append(
                    {
                        "name": name,
                        "text": _shape_text(shape),
                        "slide_order": order,
                        "has_math": shape.find(".//m:oMath", NS) is not None,
                    }
                )
            pages[page] = rows
        return len(slide_files), pages


def _manifest(run: Path, page: int) -> dict[str, Any]:
    path = run / "pages" / f"page_{page:03d}" / "manifest.json"
    payload = _load_json(path)
    if not isinstance(payload, dict):
        raise RecipeError(f"page {page} manifest must be an object")
    return payload


def _manifest_inventory(manifest: dict[str, Any], page: int) -> tuple[set[str], set[str], float, float]:
    try:
        width = _finite(manifest["source"]["width_px"], f"page {page} source.width_px")
        height = _finite(manifest["source"]["height_px"], f"page {page} source.height_px")
    except (KeyError, TypeError) as exc:
        raise RecipeError(f"page {page} manifest has no valid source dimensions") from exc
    try:
        content_box = manifest["content_box"]
        content_values = {
            key: _finite(content_box[key], f"page {page} content_box.{key}")
            for key in ("left", "top", "width", "height")
        }
        if min(content_values["left"], content_values["top"]) < 0 or min(content_values["width"], content_values["height"]) <= 0:
            raise RecipeError(f"page {page} content_box has invalid geometry")
    except (KeyError, TypeError) as exc:
        raise RecipeError(f"page {page} manifest has no valid content_box") from exc
    text_ids = {
        str(row["id"])
        for row in manifest.get("text_boxes", [])
        if isinstance(row, dict) and row.get("id")
    }
    formula_ids = {
        str(row["id"])
        for row in manifest.get("formula_inventory", [])
        if isinstance(row, dict)
        and row.get("id")
        and not str(row.get("decision", "")).startswith(("embedded-", "source-embedded-"))
    }
    overlap = text_ids & formula_ids
    if overlap:
        raise RecipeError(f"page {page} ids occur in both text and formula inventories: {sorted(overlap)}")
    return text_ids, formula_ids, width, height


def validate_recipes(
    recipes: list[dict[str, Any]],
    run: Path,
    deck: Path,
    selected_pages: set[int] | None = None,
) -> dict[str, Any]:
    slide_count, deck_shapes = read_deck_shapes(deck)
    errors: list[str] = []
    normalized: list[dict[str, Any]] = []
    used_shapes: set[tuple[int, str]] = set()
    used_ids: set[str] = set()
    manifests: dict[int, tuple[set[str], set[str], float, float]] = {}

    for line_index, line in enumerate(recipes, start=1):
        label = f"line {line_index}"
        try:
            unknown = set(line) - LINE_FIELDS
            if unknown:
                raise RecipeError(f"unknown fields: {sorted(unknown)}")
            required = {"page", "text_font_pt", "center_y_px", "start_x_px", "right_limit_px", "items"}
            missing = required - set(line)
            if missing:
                raise RecipeError(f"missing required fields: {sorted(missing)}")
            page = _positive_int(line["page"], "page")
            if page > slide_count:
                raise RecipeError(f"page {page} exceeds deck slide count {slide_count}")
            line_id = line.get("id", f"line-{line_index:04d}")
            if not isinstance(line_id, str) or not line_id:
                raise RecipeError("id must be a non-empty string")
            if line_id in used_ids:
                raise RecipeError(f"duplicate line id {line_id!r}")
            used_ids.add(line_id)

            text_font_pt = _finite(line["text_font_pt"], "text_font_pt")
            if not 4 <= text_font_pt <= 60:
                raise RecipeError("text_font_pt must be between 4 and 60")
            center_y = _finite(line["center_y_px"], "center_y_px")
            start_x = _finite(line["start_x_px"], "start_x_px")
            right_limit = _finite(line["right_limit_px"], "right_limit_px")
            items = line["items"]
            if not isinstance(items, list) or not items or not all(isinstance(item, dict) for item in items):
                raise RecipeError("items must be a non-empty array of objects")

            if page not in manifests:
                manifests[page] = _manifest_inventory(_manifest(run, page), page)
            text_ids, formula_ids, source_width, source_height = manifests[page]
            if not 0 <= center_y <= source_height:
                raise RecipeError(f"center_y_px must be within 0..{source_height:g}")
            if not 0 <= start_x < right_limit <= source_width:
                raise RecipeError(
                    f"require 0 <= start_x_px < right_limit_px <= {source_width:g}"
                )

            shapes = deck_shapes.get(page, [])
            resolved_items: list[dict[str, Any]] = []
            for item_index, item in enumerate(items, start=1):
                item_label = f"item {item_index}"
                unknown_item = set(item) - ITEM_FIELDS
                if unknown_item:
                    raise RecipeError(f"{item_label} unknown fields: {sorted(unknown_item)}")
                has_shape = "shape" in item
                has_text = "text" in item
                if has_shape == has_text:
                    raise RecipeError(f"{item_label} requires exactly one of shape or text")
                if "match_index" in item and not has_text:
                    raise RecipeError(f"{item_label} match_index is valid only with text")
                if has_shape:
                    selector = item["shape"]
                    if not isinstance(selector, str) or not selector:
                        raise RecipeError(f"{item_label} shape must be a non-empty string")
                    name_matches = [shape for shape in shapes if shape["name"] == selector]
                    if not name_matches:
                        raise RecipeError(f"{item_label} shape {selector!r} is missing from slide {page}")
                    if len(name_matches) != 1:
                        raise RecipeError(
                            f"{item_label} shape selector {selector!r} is ambiguous: {len(name_matches)} matches"
                        )
                    selected = name_matches[0]
                else:
                    selector = item["text"]
                    if not isinstance(selector, str):
                        raise RecipeError(f"{item_label} text must be a string")
                    # Native PowerPoint prose names (for example "TextBox 20")
                    # need not match manifest text-box ids. Resolve exact prose
                    # text directly in slide order; OMML shapes are excluded.
                    matches = [
                        shape for shape in shapes
                        if not shape["has_math"] and shape["text"] == selector
                    ]
                    if "match_index" in item:
                        match_index = item["match_index"]
                        if isinstance(match_index, bool) or not isinstance(match_index, int) or match_index < 0:
                            raise RecipeError(f"{item_label} match_index must be a zero-based non-negative integer")
                        if match_index >= len(matches):
                            raise RecipeError(
                                f"{item_label} match_index {match_index} exceeds {len(matches)} exact text matches"
                            )
                        selected = matches[match_index]
                    else:
                        if len(matches) != 1:
                            raise RecipeError(
                                f"{item_label} exact text selector is ambiguous: {len(matches)} matches; add match_index"
                            )
                        selected = matches[0]

                shape_name = selected["name"]
                if selected["has_math"]:
                    kind = "math"
                else:
                    kind = "prose"

                key = (page, shape_name)
                if key in used_shapes:
                    raise RecipeError(f"shape {shape_name!r} on page {page} is repeated in line recipes")
                used_shapes.add(key)

                if "expected_text" in item and "exact_expected" in item:
                    raise RecipeError(f"{item_label} cannot contain both expected_text and exact_expected")
                expected = item.get("expected_text", item.get("exact_expected"))
                if expected is not None:
                    if not isinstance(expected, str):
                        raise RecipeError(f"{item_label} exact expected text must be a string")
                    if selected["text"] != expected:
                        raise RecipeError(
                            f"{item_label} exact expected text differs for {page}/{shape_name}"
                        )

                gap = _finite(item.get("gap_before_px", 0), f"{item_label} gap_before_px")
                if not -2 <= gap <= 150:
                    raise RecipeError(f"{item_label} gap_before_px must be between -2 and 150")
                center_offset = _finite(
                    item.get("center_offset_px", 0), f"{item_label} center_offset_px"
                )
                if abs(center_offset) > 80:
                    raise RecipeError(f"{item_label} center_offset_px exceeds 80")
                requested_font = None
                if "font_pt" in item:
                    requested_font = _finite(item["font_pt"], f"{item_label} font_pt")
                    if not 4 <= requested_font <= 60:
                        raise RecipeError(f"{item_label} font_pt must be between 4 and 60")
                resolved_items.append(
                    {
                        "shape": shape_name,
                        "kind": kind,
                        "text": selected["text"],
                        "slide_order": selected["slide_order"],
                        "expected_text": expected,
                        # Per-item font_pt is a prose-only override. It is kept
                        # out of native-math plans so formula size is preserved.
                        "prose_font_pt": requested_font if kind == "prose" and requested_font is not None else text_font_pt,
                        "ignored_math_font_pt": kind == "math" and requested_font is not None,
                        "gap_before_px": gap,
                        "center_offset_px": center_offset,
                    }
                )

            normalized.append(
                {
                    "id": line_id,
                    "page": page,
                    "text_font_pt": text_font_pt,
                    "center_y_px": center_y,
                    "start_x_px": start_x,
                    "right_limit_px": right_limit,
                    "source_size_px": [source_width, source_height],
                    "items": resolved_items,
                }
            )
        except RecipeError as exc:
            errors.append(f"{label}: {exc}")

    if errors:
        raise RecipeError("\n".join(errors))
    selected = normalized if not selected_pages else [line for line in normalized if line["page"] in selected_pages]
    if not selected:
        raise RecipeError("no validated line recipes remain after page selection")
    warnings = [
        f"line {line['id']} formula {item['shape']}: font_pt ignored to preserve native math"
        for line in selected
        for item in line["items"]
        if item.get("ignored_math_font_pt")
    ]
    return {
        "schema_version": 1,
        "status": "pass",
        "line_count": len(selected),
        "warnings": warnings,
        "lines": selected,
    }


def validate_width_measurements(
    plan: dict[str, Any],
    measurements: Any,
    tolerance_px: float = 0.1,
) -> dict[str, Any]:
    if not isinstance(plan, dict) or not isinstance(plan.get("lines"), list):
        raise RecipeError("invalid normalized plan")
    if not isinstance(measurements, list):
        raise RecipeError("measurements must be an array")
    by_id: dict[str, dict[str, Any]] = {}
    for row in measurements:
        if not isinstance(row, dict) or not isinstance(row.get("id"), str):
            raise RecipeError("every measurement needs a string id")
        if row["id"] in by_id:
            raise RecipeError(f"duplicate measurement id {row['id']!r}")
        by_id[row["id"]] = row
    expected_ids = {line["id"] for line in plan["lines"]}
    if set(by_id) != expected_ids:
        raise RecipeError(
            f"measurement ids differ from plan: expected {sorted(expected_ids)}, got {sorted(by_id)}"
        )

    results = []
    for line in plan["lines"]:
        row = by_id[line["id"]]
        measured_items = row.get("items")
        if not isinstance(measured_items, list) or len(measured_items) != len(line["items"]):
            raise RecipeError(f"line {line['id']!r} measurement item count differs from plan")
        cursor = float(line["start_x_px"])
        item_results = []
        for planned, measured in zip(line["items"], measured_items):
            if not isinstance(measured, dict) or measured.get("shape") != planned["shape"]:
                raise RecipeError(f"line {line['id']!r} measurement shape order differs from plan")
            width = _finite(measured.get("width_px"), f"line {line['id']} {planned['shape']} width_px")
            if width <= 0:
                raise RecipeError(f"line {line['id']} {planned['shape']} width_px must be positive")
            cursor += float(planned["gap_before_px"])
            right = cursor + width
            if right > float(line["right_limit_px"]) + tolerance_px:
                raise RecipeError(
                    f"line {line['id']!r} overflows at {planned['shape']!r}: "
                    f"right={right:.3f}, limit={float(line['right_limit_px']):.3f}; no resize allowed"
                )
            item_results.append(
                {"shape": planned["shape"], "left_px": cursor, "width_px": width, "right_px": right}
            )
            cursor = right
        results.append({"id": line["id"], "right_px": cursor, "items": item_results})
    return {"schema_version": 1, "status": "pass", "lines": results}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--lines", type=Path)
    parser.add_argument("--run", type=Path)
    parser.add_argument("--deck", type=Path)
    parser.add_argument("--page", type=int, action="append")
    parser.add_argument("--plan", type=Path)
    parser.add_argument("--measurements", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    try:
        if args.measurements is not None or args.plan is not None:
            if args.plan is None or args.measurements is None:
                raise RecipeError("--plan and --measurements are required together")
            report = validate_width_measurements(_load_json(args.plan), _load_json(args.measurements))
            message = f"mixed-line widths passed: {len(report['lines'])} lines"
        else:
            if args.lines is None or args.run is None or args.deck is None:
                raise RecipeError("--lines, --run, and --deck are required")
            selected_pages = None
            if args.page:
                selected_pages = {_positive_int(page, "--page") for page in args.page}
            report = validate_recipes(
                load_recipe_document(args.lines), args.run, args.deck, selected_pages
            )
            message = f"mixed-line recipe passed: {report['line_count']} lines"
            for warning in report.get("warnings", []):
                print(f"warning: {warning}", file=sys.stderr)
        if args.output:
            _write_new_json(args.output, report)
        print(message)
        return 0
    except RecipeError as exc:
        print(f"mixed-line preflight failed:\n{exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())

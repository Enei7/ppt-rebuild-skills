"""Resolve formula source evidence separately from native placement geometry."""

from __future__ import annotations

import math
from dataclasses import dataclass
from pathlib import PurePosixPath
from typing import Any, Iterable


class FormulaSourceGeometryError(ValueError):
    pass


@dataclass(frozen=True)
class FormulaSourceGeometry:
    target_box_px: tuple[float, float, float, float]
    source_box_px: tuple[float, float, float, float]
    source_kind: str
    evidence: tuple[str, ...]

    def as_lists(self) -> tuple[list[float], list[float]]:
        return list(self.target_box_px), list(self.source_box_px)


def _box(value: Any, label: str) -> tuple[float, float, float, float]:
    if (
        not isinstance(value, (list, tuple))
        or len(value) != 4
        or any(isinstance(item, bool) or not isinstance(item, (int, float)) for item in value)
    ):
        raise FormulaSourceGeometryError(f"{label} must be [x, y, width, height]")
    result = tuple(float(item) for item in value)
    if not all(math.isfinite(item) for item in result) or min(result[0], result[1]) < 0 or min(result[2], result[3]) <= 0:
        raise FormulaSourceGeometryError(f"{label} is invalid: {value!r}")
    return result


def _asset_path(value: Any) -> str | None:
    if not isinstance(value, str) or not value.strip():
        return None
    return str(PurePosixPath(value.replace("\\", "/")))


def _matching_records(
    records: Iterable[Any],
    paths: set[str],
    ids: set[str],
) -> Iterable[tuple[dict[str, Any], str]]:
    for index, record in enumerate(records):
        if not isinstance(record, dict):
            continue
        path = _asset_path(record.get("path"))
        ident = record.get("id")
        if (path and path in paths) or (isinstance(ident, str) and ident in ids):
            yield record, str(index)


def resolve_formula_source_geometry(
    manifest: dict[str, Any],
    formula: dict[str, Any],
    page: int,
    source_size_px: tuple[int, int] | list[int] | None = None,
) -> FormulaSourceGeometry:
    """Resolve actual source crop and independent target placement.

    Exact source-crop evidence is joined through the formula image path and its
    image/visual/provenance records.  If no exact-crop evidence exists, the
    legacy contract uses formula_inventory.box_px for both source and target.
    Conflicting or incomplete exact-crop evidence fails instead of guessing.
    """

    ident = str(formula.get("id") or "<unnamed>")
    target = _box(formula.get("box_px"), f"formula {ident} target box_px")
    ids = {ident}
    replace_id = formula.get("replace_image_id")
    if isinstance(replace_id, str) and replace_id:
        ids.add(replace_id)

    paths: set[str] = set()
    formula_image = _asset_path(formula.get("image"))
    if formula_image:
        paths.add(formula_image)
    matched_images: list[tuple[dict[str, Any], str]] = []
    for image, index in _matching_records(manifest.get("images", []), paths, ids):
        matched_images.append((image, index))
        path = _asset_path(image.get("path"))
        if path:
            paths.add(path)
    # A path discovered through an image id may reveal additional image records.
    matched_images = list(_matching_records(manifest.get("images", []), paths, ids))

    sources: list[tuple[str, dict[str, Any]]] = []
    sources.extend((f"images[{index}]", record) for record, index in matched_images)
    sources.extend(
        (f"asset_provenance[{index}]", record)
        for record, index in _matching_records(manifest.get("asset_provenance", []), paths, ids)
    )
    sources.extend(
        (f"visual_inventory[{index}]", record)
        for record, index in _matching_records(manifest.get("visual_inventory", []), paths, ids)
    )

    crop_records: list[tuple[str, tuple[float, float, float, float]]] = []
    for label, record in sources:
        declares_crop = any(
            key in record for key in ("exact_source_crop", "source_crop_px", "source_page")
        )
        if not declares_crop:
            continue
        if record.get("exact_source_crop") is not True:
            raise FormulaSourceGeometryError(
                f"page {page} formula {ident}: {label} has source-crop metadata without exact_source_crop=true"
            )
        if "source_crop_px" not in record or "source_page" not in record:
            raise FormulaSourceGeometryError(
                f"page {page} formula {ident}: {label} has incomplete exact source-crop metadata"
            )
        source_page = record.get("source_page")
        if isinstance(source_page, bool) or not isinstance(source_page, int) or source_page != page:
            raise FormulaSourceGeometryError(
                f"page {page} formula {ident}: {label} source_page={source_page!r} does not match"
            )
        source_path = record.get("source")
        if source_path is not None:
            normalized_source = _asset_path(source_path)
            if normalized_source is None or PurePosixPath(normalized_source).name != "source.png":
                raise FormulaSourceGeometryError(
                    f"page {page} formula {ident}: {label} source must resolve to this page's source.png"
                )
        crop_records.append(
            (label, _box(record["source_crop_px"], f"{label}.source_crop_px"))
        )

    if crop_records:
        distinct = {crop for _, crop in crop_records}
        if len(distinct) != 1:
            details = ", ".join(f"{label}={list(crop)}" for label, crop in crop_records)
            raise FormulaSourceGeometryError(
                f"page {page} formula {ident}: ambiguous exact source crops: {details}"
            )
        source = crop_records[0][1]
        kind = "exact-source-crop"
        evidence = tuple(label for label, _ in crop_records)
    else:
        source = target
        kind = "formula-box-fallback"
        evidence = ("formula_inventory.box_px",)

    if source_size_px is not None:
        if len(source_size_px) != 2:
            raise FormulaSourceGeometryError("source_size_px must be [width, height]")
        width, height = map(float, source_size_px)
        if source[0] + source[2] > width or source[1] + source[3] > height:
            raise FormulaSourceGeometryError(
                f"page {page} formula {ident}: source crop {list(source)} exceeds source image {list(source_size_px)}"
            )

    return FormulaSourceGeometry(target, source, kind, evidence)


def translate_source_ink_to_target(
    source_ink_in_box: list[int] | tuple[int, int, int, int] | None,
    source_box_px: tuple[float, float, float, float] | list[float],
    target_box_px: tuple[float, float, float, float] | list[float],
) -> list[float] | None:
    """Translate source-relative ink to target-relative expected ink.

    The source glyph dimensions are preserved.  Only the difference between the
    source-crop center and target-placement center is applied; no target-box
    scaling can shrink or clip the source evidence.
    """

    if source_ink_in_box is None:
        return None
    source = _box(source_box_px, "source_box_px")
    target = _box(target_box_px, "target_box_px")
    if len(source_ink_in_box) != 4:
        raise FormulaSourceGeometryError("source ink must have four coordinates")
    dx = (target[2] - source[2]) / 2.0
    dy = (target[3] - source[3]) / 2.0
    return [float(source_ink_in_box[0]) + dx, float(source_ink_in_box[1]) + dy,
            float(source_ink_in_box[2]) + dx, float(source_ink_in_box[3]) + dy]

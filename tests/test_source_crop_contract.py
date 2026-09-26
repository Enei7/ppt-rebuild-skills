"""Smoke check for the documented exact-PDF-crop exception."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "cli" / "editppt" / "runtime"))

from validate_pptx import foreground_asset_contract_violations, page_contract_violations
from formula_renderer import formula_image_fragment


def main():
    provenance = {
        "path": "assets/qr.png",
        "source": "source.png",
        "source_type": "user-provided",
        "exact_source_crop": True,
        "source_page": 3,
        "source_crop_px": [100, 120, 80, 80],
        "provenance_note": "Exact QR crop from PDF page 3",
    }
    manifest = {
        "visual_inventory": [{"role": "foreground", "object_type": "qr", "path": "assets/qr.png", "source_type": "user-provided"}],
        "asset_provenance": [provenance],
    }
    assert not foreground_asset_contract_violations(manifest)
    provenance.pop("source_crop_px")
    assert foreground_asset_contract_violations(manifest)
    fragment = formula_image_fragment(
        formula_id="f1", image_path="assets/f1.svg", tex_source="assets/f1.tex", box_px=[10, 20, 80, 40]
    )
    assert fragment["formula_inventory"][0]["box_px"] == [10.0, 20.0, 80.0, 40.0]
    formula = {"id": "f1", "image": "assets/f1.png", "decision": "source-exact-formula-crop"}
    assert page_contract_violations({"formula_inventory": [formula], "images": []})
    formula["decision"] = "embedded-editable-text"
    assert not page_contract_violations({"formula_inventory": [formula], "images": []})


if __name__ == "__main__":
    main()

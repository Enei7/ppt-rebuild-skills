"""Read-only final-deck checks against a reconstruction run's page inventories.

This is a structural gate, not a substitute for PowerPoint render/visual review.
"""

import argparse
from collections import Counter
import json
from pathlib import Path
import posixpath
import zipfile

from lxml import etree

from equationize import NS, R, tag
from measure_math_geometry import native_formula

REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def ordered_slides(archive):
    """Follow presentation order, not ZIP member or slide filename order."""
    presentation = etree.fromstring(archive.read("ppt/presentation.xml"))
    rels = etree.fromstring(archive.read("ppt/_rels/presentation.xml.rels"))
    targets = {}
    for rel in rels.findall(f"{{{REL}}}Relationship"):
        if rel.get("TargetMode") == "External":
            continue
        target = rel.get("Target", "")
        targets[rel.get("Id")] = (target.lstrip("/") if target.startswith("/")
                                  else posixpath.normpath(posixpath.join("ppt", target)))
    return [targets[item.get(tag(R, "id"))]
            for item in presentation.xpath("./p:sldIdLst/p:sldId", namespaces=NS)]


def audit(run, deck):
    run = Path(run)
    metadata = json.loads((run / "deck_manifest.json").read_text(encoding="utf-8-sig"))
    pages = metadata["pages"]
    issues, details = [], []
    with zipfile.ZipFile(deck) as archive:
        corrupt = archive.testzip()
        if corrupt:
            raise ValueError(f"Corrupt ZIP member: {corrupt}")
        slides = ordered_slides(archive)
        if len(slides) != len(pages):
            issues.append(f"Slide count {len(slides)} differs from expected {len(pages)}")
        for index, (page, slide_path) in enumerate(zip(pages, slides), 1):
            path = run / page.get("manifest", f"{page['page_dir']}/manifest.json")
            manifest = json.loads(path.read_text(encoding="utf-8-sig"))
            formulas = [f for f in manifest.get("formula_inventory", []) if native_formula(f)]
            expected = Counter(f["id"] for f in formulas)
            root = etree.fromstring(archive.read(slide_path))
            actual = Counter()
            for shape in root.xpath(".//p:sp[.//m:oMath]", namespaces=NS):
                name = shape.xpath("string(./p:nvSpPr/p:cNvPr/@name)", namespaces=NS)
                actual[name] += len(shape.xpath(".//m:oMath", namespaces=NS))
            for name, count in expected.items():
                if count != 1:
                    issues.append(f"Page {index}: duplicate inventory id {name}")
                if actual[name] != count:
                    issues.append(f"Page {index}: {name} has {actual[name]} equations; expected {count}")
            for name in actual.keys() - expected.keys():
                issues.append(f"Page {index}: unexpected native equation {name}")
            total = len(root.xpath(".//m:oMath", namespaces=NS))
            if total != sum(expected.values()):
                issues.append(f"Page {index}: equation count {total}; expected {sum(expected.values())}")
            empty = root.xpath(".//m:nary/m:e[not(*) and not(normalize-space())]", namespaces=NS)
            if empty:
                issues.append(f"Page {index}: {len(empty)} empty n-ary bodies")
            replaced_ids = {f.get("replace_image_id", f["id"]) for f in formulas}
            replaced_paths = {f.get("image") for f in formulas if f.get("image")}
            images = manifest.get("images", [])
            expected_pictures = sum(image.get("id") not in replaced_ids and image.get("path") not in replaced_paths
                                    for image in images)
            pictures = root.xpath(".//p:pic", namespaces=NS)
            if len(pictures) != expected_pictures:
                issues.append(f"Page {index}: {len(pictures)} pictures; expected {expected_pictures} after formula replacement")
            details.append({"page": index, "equations": total, "pictures": len(pictures)})
    return {"passed": not issues, "deck": str(Path(deck).resolve()), "pages": len(slides),
            "expected_pages": len(pages), "equations": sum(p["equations"] for p in details),
            "pictures": sum(p["pictures"] for p in details), "issues": issues, "page_counts": details,
            "scope": "Structural inventory check only; visual and editor checks remain required."}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--run", required=True, type=Path)
    parser.add_argument("--deck", required=True, type=Path)
    parser.add_argument("--report", type=Path)
    args = parser.parse_args()
    result = audit(args.run, args.deck)
    payload = json.dumps(result, ensure_ascii=False, indent=2)
    if args.report:
        # Keep older evidence; callers choose a fresh report for each candidate.
        with args.report.open("x", encoding="utf-8") as stream:
            stream.write(payload + "\n")
    print(payload)
    return 0 if result["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())

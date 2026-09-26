"""Exercise inventory reconciliation, including reordered slide parts."""

import json
from pathlib import Path
import sys
import tempfile
import zipfile

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from audit_deck import audit
from equationize import P, A, M, R


def slide(name, picture=False, empty=False):
    nary = '<m:nary><m:e/></m:nary>' if empty else '<m:r><m:t>x</m:t></m:r>'
    shape = f'<p:sp><p:nvSpPr><p:cNvPr id="2" name="{name}"/></p:nvSpPr><m:oMath>{nary}</m:oMath></p:sp>'
    return f'<p:sld xmlns:p="{P}" xmlns:a="{A}" xmlns:m="{M}"><p:cSld><p:spTree>{shape}{"<p:pic/>" if picture else ""}</p:spTree></p:cSld></p:sld>'


def write_deck(path, names=("f1", "f2"), picture=False, empty=False, reverse=False):
    ids = ('r2', 'r1') if reverse else ('r1', 'r2')
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("ppt/presentation.xml", f'<p:presentation xmlns:p="{P}" xmlns:r="{R}"><p:sldIdLst><p:sldId id="256" r:id="{ids[0]}"/><p:sldId id="257" r:id="{ids[1]}"/></p:sldIdLst></p:presentation>')
        z.writestr("ppt/_rels/presentation.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Target="slides/slide9.xml"/><Relationship Id="r2" Target="/ppt/slides/slide2.xml"/></Relationships>')
        z.writestr("ppt/slides/slide2.xml", slide(names[1]))
        z.writestr("ppt/slides/slide9.xml", slide(names[0], picture, empty))


def main():
    with tempfile.TemporaryDirectory() as directory:
        run = Path(directory)
        pages = []
        for number in (1, 2):
            page = run / "pages" / f"page_{number:03}"
            page.mkdir(parents=True)
            formula = {"id": f"f{number}", "image": "assets/f.png", "decision": "source-formula-crop"}
            (page / "manifest.json").write_text(json.dumps({"formula_inventory": [formula], "images": [{"id": f"f{number}", "path": "assets/f.png"}]}))
            pages.append({"page_dir": str(page.relative_to(run))})
        (run / "deck_manifest.json").write_text(json.dumps({"pages": pages}))
        deck = run / "test.pptx"
        write_deck(deck)
        assert audit(run, deck)["passed"]  # filename/ZIP order deliberately differ
        write_deck(deck, names=("f2", "f1"))
        assert not audit(run, deck)["passed"]  # same count, wrong page formula IDs
        write_deck(deck, picture=True)
        assert any("pictures" in issue for issue in audit(run, deck)["issues"])
        write_deck(deck, empty=True)
        assert any("empty n-ary" in issue for issue in audit(run, deck)["issues"])
        write_deck(deck, reverse=True)
        assert not audit(run, deck)["passed"]
        # A valid ZIP that lacks a required slide must never pass on total math alone.
        with zipfile.ZipFile(deck, "w") as z:
            z.writestr("ppt/presentation.xml", f'<p:presentation xmlns:p="{P}"><p:sldIdLst/></p:presentation>')
            z.writestr("ppt/_rels/presentation.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')
        assert not audit(run, deck)["passed"]


if __name__ == "__main__":
    main()

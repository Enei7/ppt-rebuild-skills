"""Overlines must stay above letters; preserve unrelated accents and content."""
import sys
from pathlib import Path
from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from equationize import A, M, NS, tag, find_mml2omml, formula_shape, normalize_overline_bars


def main():
    transform = etree.XSLT(etree.parse(str(find_mml2omml())))
    xfrm = etree.Element(tag(A, "xfrm"))
    etree.SubElement(xfrm, tag(A, "off"), x="0", y="0")
    etree.SubElement(xfrm, tag(A, "ext"), cx="1828800", cy="457200")
    for tex, expected, count in (
        (r"\overline V", "V", 1),
        (r"\overline{xy}", "xy", 1),
        (r"\overline{\overline{z}}", "z", 2),
        (r"\bar{z}", "z", 1),
    ):
        shape = formula_shape({"id": "bar", "latex": tex, "box_px": [0, 0, 150, 60]}, xfrm, 2, transform)
        bars = shape.xpath(".//m:bar[m:barPr/m:pos[@m:val='top']]", namespaces=NS)
        assert len(bars) == count
        assert ''.join(shape.xpath('.//m:t/text()', namespaces=NS)) == expected
        assert not shape.xpath('.//m:acc', namespaces=NS)
        assert normalize_overline_bars(shape) == 0

    root = etree.fromstring(f'<m:oMath xmlns:m="{M}"><m:acc><m:accPr><m:chr m:val="^"/></m:accPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:acc></m:oMath>')
    original = etree.tostring(root)
    assert normalize_overline_bars(root) == 0
    assert etree.tostring(root) == original


if __name__ == '__main__':
    main()

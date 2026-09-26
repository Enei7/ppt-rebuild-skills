"""Target positive TeX spacing preservation and n-ary operand recovery."""

import sys
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from equationize import A, M, NS, MATHML, fill_empty_nary_bases, formula_shape, preserve_positive_mathml_spacing, tag, find_mml2omml


def xfrm():
    value = etree.Element(tag(A, "xfrm"))
    etree.SubElement(value, tag(A, "off"), x="0", y="0")
    etree.SubElement(value, tag(A, "ext"), cx="1828800", cy="457200")
    return value


def main():
    synthetic = etree.fromstring(
        f'<math xmlns="{MATHML}"><mrow>'
        '<mspace width="0.167em"/><mspace width="0.278em"/>'
        '<mspace width="1em"/><mspace width="2em"/>'
        '<mspace width="0em"/><mspace width="-1em"/><mspace width="0.5em"/>'
        '</mrow></math>'
    )
    assert preserve_positive_mathml_spacing(synthetic) == 4
    assert synthetic.xpath(".//mml:mtext/text()", namespaces={"mml": MATHML}) == [
        "\u2009", "\u2005", "\u2003", "\u2003\u2003"
    ]
    assert synthetic.xpath(".//mml:mspace/@width", namespaces={"mml": MATHML}) == [
        "0em", "-1em", "0.5em"
    ]

    transform = etree.XSLT(etree.parse(str(find_mml2omml())))
    spaced = formula_shape(
        {"id": "ordinary-spacing", "latex": r"0\,x\;y\quad\text{or}\qquad(D)", "box_px": [0, 0, 400, 60]},
        xfrm(), 2, transform,
    )
    text = "".join(spaced.xpath(".//m:t/text()", namespaces=NS))
    assert "\u2009" in text
    assert "\u2005" in text
    assert "\u2003or\u2003\u2003" in text

    nary = formula_shape(
        {"id": "nary-spacing", "latex": r"\sum_{m=1}^{\infty}\quad x_m", "box_px": [0, 0, 300, 90]},
        xfrm(), 3, transform,
    )
    bodies = nary.xpath(".//m:nary/m:e", namespaces=NS)
    assert len(bodies) == 1
    assert "".join(bodies[0].itertext()).strip() == "xm"
    assert bodies[0].xpath(".//m:t[text()='\u2003']", namespaces=NS)

    manual = etree.fromstring(
        f'<m:oMath xmlns:m="{M}"><m:nary><m:e/></m:nary>'
        '<m:r><m:t>\u2003</m:t></m:r><m:r><m:t>x</m:t></m:r></m:oMath>'
    )
    assert fill_empty_nary_bases(manual) == 1
    body = manual.find(f".//{tag(M, 'nary')}/{tag(M, 'e')}")
    assert "".join(body.itertext()).strip() == "x"
    assert body.find(f"./{tag(M, 'r')}/{tag(M, 't')}").text == "\u2003"


if __name__ == "__main__":
    main()

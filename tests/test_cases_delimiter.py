"""Piecewise cases keep a native spanning brace and no closing delimiter."""
import sys
from pathlib import Path
from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from equationize import A, M, NS, tag, find_mml2omml, formula_shape


def main():
    transform = etree.XSLT(etree.parse(str(find_mml2omml())))
    xfrm = etree.Element(tag(A, 'xfrm'))
    etree.SubElement(xfrm, tag(A, 'off'), x='0', y='0')
    etree.SubElement(xfrm, tag(A, 'ext'), cx='3657600', cy='1371600')
    for count, tex in [
        (2, r'f(x)=\begin{cases}1,&x>0\\0,&x\leq0\end{cases}'),
        (3, r'\begin{cases}0,&m\ne n\\1/2,&m=n>0\\1,&m=n=0\end{cases}'),
    ]:
        shape = formula_shape({'id':'cases','latex':tex,'box_px':[0,0,400,150]}, xfrm, 2, transform)
        delimiters = shape.xpath('.//m:d[m:dPr/m:begChr[@m:val="{"]]', namespaces=NS)
        assert len(delimiters) == 1
        assert delimiters[0].find('m:dPr/m:endChr', NS).get(tag(M,'val')) == ''
        assert len(delimiters[0].xpath('.//m:m/m:mr', namespaces=NS)) == count
        assert not shape.xpath('.//m:r/m:t[text()="{"]', namespaces=NS)
    ordinary = formula_shape({'id':'ordinary','latex':r'\{x,y\}','box_px':[0,0,200,60]}, xfrm, 2, transform)
    assert not ordinary.xpath('.//m:m', namespaces=NS)


if __name__ == '__main__':
    main()

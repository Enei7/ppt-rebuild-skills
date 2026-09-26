"""Replace editppt formula pictures with native Office equations."""

import argparse
import copy
import hashlib
import json
import os
import posixpath
import re
import tempfile
import zipfile
from pathlib import Path

try:
    import latex2mathml.converter
    from lxml import etree
except ImportError as exc:
    raise SystemExit("Install dependencies with: python -m pip install lxml latex2mathml") from exc


P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
NS = {"p": P, "a": A, "r": R, "m": M}
EMU = 914400
SLIDE_RE = re.compile(r"ppt/slides/slide(\d+)\.xml$")
NARY_BOUNDARIES = {"=", "+", "-", "−", ",", ";", ")", "]", "<", ">", "≤", "≥"}
FUNCTION_NAMES = {"sin", "cos", "tan", "cot", "sec", "csc", "sinh", "cosh", "tanh", "ln", "log", "exp", "arcsin", "arccos", "arctan", "arccot"}


def tag(namespace, name):
    return f"{{{namespace}}}{name}"


def find_mml2omml(explicit=None):
    if explicit:
        candidates = [Path(explicit).expanduser()]
    else:
        candidates = []
        if os.environ.get("MML2OMML_XSL"):
            candidates.append(Path(os.environ["MML2OMML_XSL"]).expanduser())
        for variable in ("ProgramFiles", "ProgramFiles(x86)"):
            root = os.environ.get(variable)
            if root:
                office = Path(root) / "Microsoft Office"
                candidates.extend(sorted(office.glob("root/Office*/MML2OMML.XSL")))
                candidates.extend(sorted(office.glob("Office*/MML2OMML.XSL")))
        candidates.extend(
            [
                Path("/Applications/Microsoft PowerPoint.app/Contents/Resources/MML2OMML.XSL"),
                Path("/Applications/Microsoft Word.app/Contents/Resources/MML2OMML.XSL"),
            ]
        )
    for candidate in candidates:
        if candidate.is_file():
            return candidate.resolve()
    raise FileNotFoundError(
        "MML2OMML.XSL was not found; pass --mml2omml or set MML2OMML_XSL"
    )


def fill_empty_nary_bases(root):
    """MML2OMML leaves sum/integral bodies empty, which Office draws as dotted boxes."""
    repaired = 0
    for nary in reversed(root.findall(f".//{tag(M, 'nary')}")):
        base = nary.find(tag(M, "e"))
        if base is None or len(base) or (base.text or "").strip():
            continue
        next_item = nary.getnext()
        next_text = "" if next_item is None else "".join(next_item.itertext()).strip()
        if next_item is not None and next_item.tag.startswith(f"{{{M}}}") and next_text not in NARY_BOUNDARIES:
            base.append(next_item)
        else:
            # Standalone operators (e.g. a sum symbol embedded in a sentence)
            # have no operand in this equation shape. A zero-width body keeps
            # their native limits without displaying an Office placeholder.
            run = etree.SubElement(base, tag(M, "r"))
            etree.SubElement(run, tag(M, "t")).text = "\u2060"
        repaired += 1
    return repaired


def space_function_runs(root):
    """Office's MathML transform drops mspace, so keep function gaps in OMML."""
    for run in list(root.findall(f".//{tag(M, 'r')}")):
        if "".join(run.itertext()) not in FUNCTION_NAMES:
            continue
        atom = run
        container = atom.getparent()
        if container.tag == tag(M, "e") and len(container) == 1 and container.getparent().tag in {tag(M, "sSup"), tag(M, "sSub"), tag(M, "sSubSup")}:
            atom = container.getparent()
        parent = atom.getparent()
        index = parent.index(atom)
        for offset, neighbor in ((0, parent[index - 1] if index else None), (1, parent[index + 1] if index + 1 < len(parent) else None)):
            if neighbor is None or neighbor.tag not in {tag(M, name) for name in ("r", "sSup", "sSub", "sSubSup", "f", "rad")}:
                continue
            text = "".join(neighbor.itertext())
            if not text or not (text[-1].isalnum() if offset == 0 else text[0].isalnum()):
                continue
            gap = etree.Element(tag(M, "r"))
            etree.SubElement(gap, tag(M, "t")).text = "\u2009"
            parent.insert(parent.index(atom) + offset, gap)


def formula_shape(formula, xfrm, shape_id, transform):
    tex = formula["latex"]
    # Bare integral tokens otherwise become small ordinary text in Office's XSL.
    tex = re.sub(r"\\int(?![A-Za-z]|\s*[_^]|\\limits)", r"\\int_{}^{}", tex)
    for style in (r"\displaystyle", r"\textstyle", r"\scriptstyle", r"\scriptscriptstyle"):
        tex = tex.replace(style, "")
    if r"\begin{aligned}" in tex:
        # ponytail: latex2mathml mishandles aligned '&'; matrix keeps the line breaks.
        tex = tex.replace(r"\begin{aligned}", r"\begin{matrix}")
        tex = tex.replace(r"\end{aligned}", r"\end{matrix}").replace("&", "")
    mathml = etree.fromstring(
        latex2mathml.converter.convert(tex.replace(r"\arg", r"\operatorname{arg}")).encode()
    )
    unknown = [node.text for node in mathml.iter() if node.text and node.text.startswith("\\")]
    if unknown:
        raise ValueError(f"Unconverted LaTeX command in {formula['id']}: {unknown}")
    # The Office stylesheet understands mfenced, but loses stretchy mo fences.
    mathml_ns = "http://www.w3.org/1998/Math/MathML"
    for node in mathml.iter():
        if node.text in FUNCTION_NAMES and node.tag in {tag(mathml_ns, "mi"), tag(mathml_ns, "mo")}:
            node.tag = tag(mathml_ns, "mi")
            node.set("mathvariant", "normal")
    for row in reversed(list(mathml.iter(tag(mathml_ns, "mrow")))):
        children = list(row)
        if len(children) < 3:
            continue
        left, right = children[0], children[-1]
        if (left.get("fence") == right.get("fence") == "true"
                and left.get("form") == "prefix" and right.get("form") == "postfix"
                and left.get("stretchy") == right.get("stretchy") == "true"):
            fenced = etree.Element(tag(mathml_ns, "mfenced"), open=left.text or "", close=right.text or "", separators="")
            content = etree.SubElement(fenced, tag(mathml_ns, "mrow"))
            for child in children[1:-1]:
                content.append(child)
            row.getparent().replace(row, fenced)
    # Matrix environments emit plain mo siblings rather than stretchy fences.
    # Office otherwise renders parentheses only as small middle-row characters.
    matrix_fences = {"(": ")", "[": "]", "{": "}", "|": "|", "‖": "‖"}
    for table in reversed(list(mathml.iter(tag(mathml_ns, "mtable")))):
        parent = table.getparent()
        index = parent.index(table)
        if index == 0 or index + 1 >= len(parent):
            continue
        left, right = parent[index - 1], parent[index + 1]
        if (left.tag == right.tag == tag(mathml_ns, "mo")
                and left.text in matrix_fences and matrix_fences[left.text] == right.text):
            fenced = etree.Element(tag(mathml_ns, "mfenced"), open=left.text, close=right.text, separators="")
            left.addprevious(fenced)
            parent.remove(left)
            parent.remove(right)
            fenced.append(table)
    # Office can split astral bold-letter surrogate pairs during font-size edits.
    # Express the same typography as BMP letters plus MathML style instead.
    for node in mathml.iter():
        if node.text and len(node.text) == 1:
            code = ord(node.text)
            for start, base in ((0x1D400, ord("A")), (0x1D41A, ord("a"))):
                if start <= code < start + 26:
                    node.text = chr(base + code - start)
                    node.set("mathvariant", "bold")
                    break
    result = transform(mathml)
    omath = result.find(f".//{tag(M, 'oMath')}")
    if omath is None and result.getroot().tag == tag(M, "oMath"):
        omath = result.getroot()
    if omath is None:
        raise ValueError(f"MathML conversion produced no equation: {tex}")
    # Preserve text-style fractions as a smaller math argument, not a tall display.
    for fraction in list(omath.findall(f".//{tag(M, 'f')}")):
        levels = fraction.xpath("./m:num/m:argPr/m:scrLvl | ./m:den/m:argPr/m:scrLvl", namespaces=NS)
        if len(levels) == 2 and all(level.get(tag(M, "val")) == "0" for level in levels):
            for level in levels:
                level.getparent().remove(level)
            box = etree.Element(tag(M, "box"))
            base = etree.SubElement(box, tag(M, "e"))
            props = etree.SubElement(base, tag(M, "argPr"))
            etree.SubElement(props, tag(M, "argSz"), {tag(M, "val"): "-1"})
            fraction.getparent().replace(fraction, box)
            base.append(fraction)
    fill_empty_nary_bases(omath)
    for nary in omath.findall(f".//{tag(M, 'nary')}"):
        props = nary.find(tag(M, "naryPr"))
        if props is None:
            props = etree.Element(tag(M, "naryPr"))
            nary.insert(0, props)
        for limit in ("sub", "sup"):
            body = nary.find(tag(M, limit))
            if body is None or (not len(body) and not (body.text or "").strip()):
                hidden = props.find(tag(M, limit + "Hide"))
                if hidden is None:
                    hidden = etree.SubElement(props, tag(M, limit + "Hide"))
                hidden.set(tag(M, "val"), "1")
    space_function_runs(omath)
    if omath.xpath(".//m:nary/m:sub/m:eqArr | .//m:nary/m:sup/m:eqArr", namespaces=NS):
        raise ValueError(
            f"{formula['id']}: multiline integral/sum limit is not PowerPoint-safe; "
            "place its annotation in separate editable objects"
        )

    shape = etree.Element(tag(P, "sp"), nsmap={"a14": A14, "m": M})
    nv = etree.SubElement(shape, tag(P, "nvSpPr"))
    etree.SubElement(nv, tag(P, "cNvPr"), id=str(shape_id), name=formula["id"])
    etree.SubElement(nv, tag(P, "cNvSpPr"), txBox="1")
    etree.SubElement(nv, tag(P, "nvPr"))
    shape_props = etree.SubElement(shape, tag(P, "spPr"))
    shape_props.append(copy.deepcopy(xfrm))
    geometry = etree.SubElement(shape_props, tag(A, "prstGeom"), prst="rect")
    etree.SubElement(geometry, tag(A, "avLst"))
    etree.SubElement(shape_props, tag(A, "noFill"))
    body = etree.SubElement(shape, tag(P, "txBody"))
    etree.SubElement(body, tag(A, "bodyPr"), wrap="none")
    etree.SubElement(body, tag(A, "lstStyle"))
    paragraph = etree.SubElement(body, tag(A, "p"))
    wrapper = etree.SubElement(paragraph, tag(A14, "m"))
    wrapper.append(copy.deepcopy(omath))
    size = min(24, max(10, formula["box_px"][3] * 0.25))
    if r"\sum" in tex and formula["box_px"][2] < 300 and formula["box_px"][3] >= 150:
        size = min(size, 14)
    font_size = round(float(formula.get("native_font_pt", size)) * 100)
    color = formula.get("native_color_hex", "000000").lstrip("#")
    if not re.fullmatch(r"[0-9A-Fa-f]{6}", color):
        raise ValueError(f"Invalid native_color_hex in {formula['id']}: {color}")
    for run in wrapper.findall(f".//{tag(M, 'r')}"):
        run_props = etree.Element(tag(A, "rPr"), lang="en-US", sz=str(font_size))
        fill = etree.SubElement(run_props, tag(A, "solidFill"))
        etree.SubElement(fill, tag(A, "srgbClr"), val=color.upper())
        etree.SubElement(run_props, tag(A, "latin"), typeface="Cambria Math")
        run.insert(0, run_props)
    etree.SubElement(paragraph, tag(A, "endParaRPr"), lang="en-US", sz=str(font_size))
    return shape


def xfrm_from_box(formula, request):
    box = formula["box_px"]
    source = request["source_size_px"]
    content = request["content_box"]
    x = (content["left"] + box[0] / source["width"] * content["width"]) * EMU
    y = (content["top"] + box[1] / source["height"] * content["height"]) * EMU
    width = box[2] / source["width"] * content["width"] * EMU
    height = box[3] / source["height"] * content["height"] * EMU
    xfrm = etree.Element(tag(A, "xfrm"))
    etree.SubElement(xfrm, tag(A, "off"), x=str(round(x)), y=str(round(y)))
    etree.SubElement(xfrm, tag(A, "ext"), cx=str(round(width)), cy=str(round(height)))
    return xfrm


def replace_slide(archive, slide_path, page_dir, transform):
    slide_no = int(SLIDE_RE.fullmatch(slide_path).group(1))
    rels_path = f"ppt/slides/_rels/slide{slide_no}.xml.rels"
    root = etree.fromstring(archive.read(slide_path))
    if rels_path in archive.namelist():
        rels = etree.fromstring(archive.read(rels_path))
        rid_to_media = {
            rel.get("Id"): posixpath.normpath(posixpath.join("ppt/slides", rel.get("Target")))
            for rel in rels.findall(tag(REL, "Relationship"))
            if rel.get("Type", "").endswith("/image") and rel.get("TargetMode") != "External"
        }
    else:
        rid_to_media = {}
    digest_to_pics = {}
    for picture in root.xpath(".//p:pic", namespaces=NS):
        blips = picture.xpath(".//a:blip", namespaces=NS)
        media = rid_to_media.get(blips[0].get(tag(R, "embed"))) if blips else None
        if media:
            digest = hashlib.sha256(archive.read(media)).hexdigest()
            digest_to_pics.setdefault(digest, []).append(picture)

    manifest = json.loads((page_dir / "manifest.json").read_text(encoding="utf-8"))
    request = json.loads((page_dir / "page_request.json").read_text(encoding="utf-8"))
    shape_tree = root.find(f".//{tag(P, 'spTree')}")
    next_id = max((int(value) for value in root.xpath(".//p:cNvPr/@id", namespaces=NS)), default=1) + 1
    count = 0
    for formula in manifest.get("formula_inventory", []):
        if not isinstance(formula, dict) or str(formula.get("decision", "")).startswith(
            ("embedded-", "source-embedded-")
        ):
            continue
        formula = dict(formula)
        if not formula.get("latex"):
            formula["latex"] = (page_dir / formula["tex_source"]).read_text(encoding="utf-8").strip()
        image_path = formula.get("image")
        if not image_path and formula.get("replace_image_id"):
            image_path = next(
                (item["path"] for item in manifest.get("images", []) if item.get("id") == formula["replace_image_id"]),
                None,
            )
            if not image_path:
                raise ValueError(f"Replacement image not declared on slide {slide_no}: {formula['id']}")
        picture = None
        if image_path:
            digest = hashlib.sha256((page_dir / image_path).read_bytes()).hexdigest()
            matches = digest_to_pics.get(digest, [])
            if not matches:
                raise ValueError(f"Formula picture not found on slide {slide_no}: {image_path}")
            picture = matches.pop(0)
        if picture is not None:
            xfrm = picture.find(f".//{tag(A, 'xfrm')}")
            parent = picture.getparent()
            index = parent.index(picture)
            parent.remove(picture)
        else:
            xfrm = xfrm_from_box(formula, request)
            candidate = next(
                (item for item in manifest.get("text_boxes", []) if item.get("id") == f"{formula['id']}_candidate"),
                None,
            )
            if candidate:
                matches = [
                    shape
                    for shape in shape_tree.findall(tag(P, "sp"))
                    if "".join(shape.xpath(".//a:t/text()", namespaces=NS)) == candidate["text"]
                ]
                if len(matches) != 1:
                    raise ValueError(f"Formula candidate text not uniquely found on slide {slide_no}: {formula['id']}")
                parent = matches[0].getparent()
                index = parent.index(matches[0])
                parent.remove(matches[0])
            else:
                parent, index = shape_tree, len(shape_tree)
        parent.insert(index, formula_shape(formula, xfrm, next_id, transform))
        next_id += 1
        count += 1
    return etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True), count


def equationize(run_dir, source_pptx, output_pptx, first_page=1, mml2omml=None):
    run_dir = Path(run_dir).resolve()
    source_pptx = Path(source_pptx).resolve()
    output_pptx = Path(output_pptx).resolve()
    if source_pptx == output_pptx:
        raise ValueError("--output must differ from --input")
    if output_pptx.exists():
        raise FileExistsError(f"Refusing to overwrite existing output: {output_pptx}")
    transform = etree.XSLT(etree.parse(str(find_mml2omml(mml2omml))))
    output_pptx.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{output_pptx.stem}-", suffix=".pptx", dir=output_pptx.parent
    )
    os.close(descriptor)
    temporary = Path(temporary_name)
    total = 0
    changed = []
    try:
        with zipfile.ZipFile(source_pptx) as source, zipfile.ZipFile(temporary, "w") as output:
            slide_paths = sorted(
                (name for name in source.namelist() if SLIDE_RE.fullmatch(name)),
                key=lambda name: int(SLIDE_RE.fullmatch(name).group(1)),
            )
            replacements = {}
            for offset, slide_path in enumerate(slide_paths):
                page_dir = run_dir / "pages" / f"page_{first_page + offset:03d}"
                manifest = json.loads((page_dir / "manifest.json").read_text(encoding="utf-8"))
                if manifest.get("formula_inventory"):
                    xml, count = replace_slide(source, slide_path, page_dir, transform)
                    replacements[slide_path] = xml
                    total += count
                    changed.append(slide_path)
            for item in source.infolist():
                payload = replacements[item.filename] if item.filename in replacements else source.read(item.filename)
                if item.filename.startswith("ppt/theme/"):
                    payload = payload.replace(b"PingFang SC", b"Microsoft YaHei")
                output.writestr(item, payload)
        with zipfile.ZipFile(temporary) as result:
            if result.testzip() is not None:
                raise ValueError("output PPTX contains a corrupt ZIP member")
            found = 0
            for name in changed:
                root = etree.fromstring(result.read(name))
                found += len(root.xpath(".//m:oMath", namespaces=NS))
                if root.xpath(".//m:nary/m:e[not(*) and not(normalize-space())]", namespaces=NS):
                    raise ValueError(f"empty Office nary body remains in {name}")
            if found < total:
                raise ValueError(f"native equation check failed: expected {total}, found {found}")
        os.replace(temporary, output_pptx)
    finally:
        temporary.unlink(missing_ok=True)
    print(f"native_equations={total} output={output_pptx}")


def repair_existing(source_pptx, output_pptx):
    source_pptx = Path(source_pptx).resolve()
    output_pptx = Path(output_pptx).resolve()
    if source_pptx == output_pptx or output_pptx.exists():
        raise ValueError("--output must be a new path, different from --input")
    output_pptx.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{output_pptx.stem}-", suffix=".pptx", dir=output_pptx.parent
    )
    os.close(descriptor)
    temporary = Path(temporary_name)
    repaired = 0
    try:
        with zipfile.ZipFile(source_pptx) as source, zipfile.ZipFile(temporary, "w") as output:
            for item in source.infolist():
                payload = source.read(item.filename)
                if SLIDE_RE.fullmatch(item.filename):
                    root = etree.fromstring(payload)
                    count = fill_empty_nary_bases(root)
                    if count:
                        repaired += count
                        payload = etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
                output.writestr(item, payload)
        with zipfile.ZipFile(temporary) as result:
            if result.testzip() is not None:
                raise ValueError("repaired PPTX contains a corrupt ZIP member")
            for name in result.namelist():
                if SLIDE_RE.fullmatch(name):
                    root = etree.fromstring(result.read(name))
                    if root.xpath(".//m:nary/m:e[not(*) and not(normalize-space())]", namespaces=NS):
                        raise ValueError(f"empty Office nary body remains in {name}")
        os.replace(temporary, output_pptx)
    finally:
        temporary.unlink(missing_ok=True)
    print(f"repaired_empty_nary={repaired} output={output_pptx}")


def audit_existing(source_pptx):
    equations = 0
    empty = 0
    with zipfile.ZipFile(source_pptx) as source:
        if source.testzip() is not None:
            raise ValueError("PPTX contains a corrupt ZIP member")
        for name in source.namelist():
            if SLIDE_RE.fullmatch(name):
                root = etree.fromstring(source.read(name))
                equations += len(root.xpath(".//m:oMath", namespaces=NS))
                empty += len(root.xpath(".//m:nary/m:e[not(*) and not(normalize-space())]", namespaces=NS))
    print(f"native_equations={equations} empty_nary={empty}")
    if empty:
        raise ValueError(f"{empty} empty Office sum/integral bodies remain")


def self_check(mml2omml=None):
    transform = etree.XSLT(etree.parse(str(find_mml2omml(mml2omml))))
    xfrm = etree.Element(tag(A, "xfrm"))
    etree.SubElement(xfrm, tag(A, "off"), x="0", y="0")
    etree.SubElement(xfrm, tag(A, "ext"), cx="914400", cy="457200")
    shape = formula_shape({"id": "self-check", "latex": r"x^2+1", "box_px": [0, 0, 100, 50]}, xfrm, 2, transform)
    assert shape.xpath("count(.//m:oMath)", namespaces=NS) == 1.0
    assert shape.xpath("string(.//a:latin/@typeface)", namespaces=NS) == "Cambria Math"
    spaced = formula_shape({"id": "function-gap", "latex": r"\sec x\tan x", "box_px": [0, 0, 200, 50]}, xfrm, 2, transform)
    assert len(spaced.xpath(".//m:t[text()='\u2009']", namespaces=NS)) == 3
    scripted = formula_shape({"id": "scripted-function-gap", "latex": r"a^x\ln N_0+\sec^2 x", "box_px": [0, 0, 200, 50]}, xfrm, 2, transform)
    assert len(scripted.xpath(".//m:t[text()='\u2009']", namespaces=NS)) == 3
    named = formula_shape({"id": "named-function-gap", "latex": r"(\operatorname{arccot} x)'", "box_px": [0, 0, 200, 50]}, xfrm, 2, transform)
    assert named.xpath(".//m:r[m:t='arccot']/m:rPr/m:sty[@m:val='p']", namespaces=NS)
    assert named.xpath(".//m:t[text()='\u2009']", namespaces=NS)
    space_function_runs(scripted)
    assert len(scripted.xpath(".//m:t[text()='\u2009']", namespaces=NS)) == 3
    bold = formula_shape({"id": "bold-vector", "latex": r"\mathbf{F}=m\mathbf{a}", "box_px": [0, 0, 235, 58]}, xfrm, 2, transform)
    assert "".join(bold.xpath(".//m:t/text()", namespaces=NS)) == "F=ma"
    assert len(bold.xpath(".//m:sty[@m:val='b']", namespaces=NS)) == 2
    colored = formula_shape({"id": "red-inline", "latex": "d/dx", "native_color_hex": "FF0000", "box_px": [0, 0, 100, 50]}, xfrm, 2, transform)
    assert set(colored.xpath(".//a:srgbClr/@val", namespaces=NS)) == {"FF0000"}
    integral = formula_shape({"id": "bare-integral", "latex": r"\int P\,dx", "box_px": [0, 0, 200, 80]}, xfrm, 2, transform)
    assert len(integral.xpath(".//m:nary", namespaces=NS)) == 1
    assert not integral.xpath(".//m:nary/m:e[not(*)]", namespaces=NS)
    assert len(integral.xpath(".//m:naryPr/*[self::m:subHide or self::m:supHide][@m:val='1']", namespaces=NS)) == 2
    fence = formula_shape({"id": "tall-fence", "latex": r"f\left(\frac{y}{x}\right)", "box_px": [0, 0, 100, 80]}, xfrm, 2, transform)
    assert len(fence.xpath(".//m:d/m:e/m:f", namespaces=NS)) == 1
    compact = formula_shape({"id": "compact-fraction", "latex": r"x=\tfrac12gt^2", "box_px": [0, 0, 200, 50]}, xfrm, 2, transform)
    assert len(compact.xpath(".//m:box/m:e[m:argPr/m:argSz[@m:val='-1']]/m:f", namespaces=NS)) == 1
    assert not compact.xpath(".//m:scrLvl", namespaces=NS)
    for environment in ("pmatrix", "bmatrix", "Bmatrix", "vmatrix", "Vmatrix"):
        matrix = formula_shape({"id": "matrix-fence", "latex": "G=\\begin{" + environment + r"}g_{11}&g_{12}\\g_{21}&g_{22}\end{" + environment + "}", "box_px": [0, 0, 300, 150]}, xfrm, 2, transform)
        assert len(matrix.xpath(".//m:d/m:e/m:m", namespaces=NS)) == 1, environment
        assert len(matrix.xpath(".//m:m/m:mr", namespaces=NS)) == 2, environment
    for latex in (r"\sum_{k=1}^{\infty} u_k(z)", r"\int_C u_k(z)\,dz", r"\sum_0^\infty"):
        sample = formula_shape({"id": "nary-check", "latex": latex, "box_px": [0, 0, 100, 50]}, xfrm, 3, transform)
        assert sample.xpath("count(.//m:nary/m:e[not(*) and not(normalize-space())])", namespaces=NS) == 0
    try:
        formula_shape({"id": "unsafe-limit", "latex": r"\int_{\substack{a\\b}}x\,dx", "box_px": [0, 0, 100, 50]}, xfrm, 4, transform)
    except ValueError as exc:
        assert "not PowerPoint-safe" in str(exc)
    else:
        raise AssertionError("multiline n-ary limit must be rejected")
    print("self-check=ok")


def parse_args():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--run-dir", type=Path)
    parser.add_argument("--input", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--first-page", type=int, default=1)
    parser.add_argument("--mml2omml", type=Path)
    parser.add_argument("--self-check", action="store_true")
    parser.add_argument("--repair-existing", action="store_true", help="remove empty n-ary placeholders in an already editable PPTX")
    parser.add_argument("--audit-existing", action="store_true", help="count native equations and reject dotted n-ary placeholders")
    args = parser.parse_args()
    if args.audit_existing and not args.input:
        parser.error("--input is required with --audit-existing")
    if args.repair_existing and not all((args.input, args.output)):
        parser.error("--input and --output are required with --repair-existing")
    if not args.self_check and not args.repair_existing and not args.audit_existing and not all((args.run_dir, args.input, args.output)):
        parser.error("--run-dir, --input, and --output are required unless --self-check is used")
    if args.first_page < 1:
        parser.error("--first-page must be at least 1")
    return args


if __name__ == "__main__":
    arguments = parse_args()
    if arguments.self_check:
        self_check(arguments.mml2omml)
    elif arguments.audit_existing:
        audit_existing(arguments.input)
    elif arguments.repair_existing:
        repair_existing(arguments.input, arguments.output)
    else:
        equationize(
            arguments.run_dir,
            arguments.input,
            arguments.output,
            arguments.first_page,
            arguments.mml2omml,
        )

"""
Shared helpers for slide builders: library prototype loading, library-line cloning,
table cell borders, connectors and SVG icon injection.
"""

import copy
import io
import os
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pptx
import pptx.enum.shapes
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Pt
from pptx.oxml import parse_xml

from ..color_tokens import hex_to_rgb

BASE_DIR = Path(__file__).resolve().parents[2]
LIBRARY_PATH = BASE_DIR / "Library_copy.pptx"
LIBRARY_INSPECT_PATH = BASE_DIR / "Library_inspect.pptx"
AFTER_PATH = BASE_DIR / "After_copy.pptx"

LIBRARY_LINE_NAMES = {
    "side_fade_r": "Grey Side Fade R",
    "centre_fade_h": "Grey Centre Fade H",
    "solid_h": "Grey Solid H",
    "solid_v": "Grey Solid V",
    "centre_fade_v": "Grey Centre Fade V",
    "white_side_fade_r": "White Side Fade R Divider",
    "straight_connector_8": "Straight Connector 8",
}


def clean_element(elem):
    """Strips any internal custDataLst / tags elements that refer to missing rId relationships."""
    for parent in elem.iter():
        for child in list(parent):
            if child.tag.endswith("custDataLst") or child.tag.endswith("tags"):
                parent.remove(child)
    return elem


def get_shape_by_name(slide, name):
    for s in slide.shapes:
        if s.name == name:
            return s
    return None


def add_library_line(slide, proto_shape, left_pt, top_pt, width_pt=0.0, height_pt=0.0):
    """Clones an authentic library line prototype, cleans it, and positions it precisely.
    Supports all 5 Library line prototypes:
      - Grey Side Fade R (horizontal, right fade)
      - Grey Centre Fade H (horizontal, center fade)
      - Grey Solid H (horizontal, solid)
      - Grey Solid V (vertical, solid)
      - Grey Centre Fade V (vertical, center fade)
    """
    elem = clean_element(copy.deepcopy(proto_shape._element))
    xfrm = elem.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm")
    if xfrm is not None:
        if "rot" in xfrm.attrib:
            del xfrm.attrib["rot"]
        off = xfrm.find("{http://schemas.openxmlformats.org/drawingml/2006/main}off")
        ext = xfrm.find("{http://schemas.openxmlformats.org/drawingml/2006/main}ext")
        if off is not None:
            off.set("x", str(int(left_pt * 12700)))
            off.set("y", str(int(top_pt * 12700)))
        if ext is not None:
            name = getattr(proto_shape, "name", "")
            is_vert = (height_pt > 0 and width_pt <= 0) or (" V" in name)
            if is_vert:
                ext.set("cx", "0")
                ext.set("cy", str(int((height_pt if height_pt > 0 else 100.0) * 12700)))
                lin = elem.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}lin")
                if lin is not None:
                    lin.set("ang", "5400000")
            else:
                ext.set("cx", str(int((width_pt if width_pt > 0 else 100.0) * 12700)))
                ext.set("cy", "0")
                lin = elem.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}lin")
                if lin is not None:
                    lin.set("ang", "0")
    slide.shapes._spTree.append(elem)
    return slide.shapes[len(slide.shapes) - 1]


def set_shape_corner_radius(shape, radius_pt=4.0):
    """Sets the border radius of any rounded rectangle shape to exact radius_pt (default 4px)."""
    try:
        prst = shape._element.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom")
        if prst is not None:
            geom = prst.get("prst", "")
            if geom == "roundRect":
                w = shape.width.pt
                h = shape.height.pt
                min_dim = min(w, h)
                if min_dim > 0:
                    shape.adjustments[0] = min(0.5, max(0.0, radius_pt / min_dim))
                    return True
            elif geom == "round2SameRect":
                w = shape.width.pt
                h = shape.height.pt
                min_dim = min(w, h)
                if min_dim > 0 and len(shape.adjustments) >= 2:
                    shape.adjustments[0] = 0.0
                    shape.adjustments[1] = min(0.5, max(0.0, radius_pt / min_dim))
                    return True
    except Exception:
        pass
    return False


def enforce_slide_corner_radii(slide, radius_pt=4.0):
    """Iterates through all shapes in a slide and enforces exact 4px corner radius on any rounded rectangle."""
    for sh in slide.shapes:
        set_shape_corner_radius(sh, radius_pt)


def enforce_deck_corner_radii(prs, radius_pt=4.0):
    """Iterates through all slides in a presentation and enforces exact 4px corner radius on all rounded boxes."""
    for slide in prs.slides:
        enforce_slide_corner_radii(slide, radius_pt)



def get_decor_shadow_proto(lib_prs):
    """Finds the authentic Decor_Box_with_Shadow_1 prototype shape in Library."""
    for s in lib_prs.slides:
        sh = get_shape_by_name(s, "Decor_Box_with_Shadow_1")
        if sh is not None:
            return sh
    return None


def apply_library_shadow(shape, proto_shape=None):
    """Applies the exact shadow settings from Decor_Box_with_Shadow_1 in Library_copy Slide 6:
    <a:outerShdw blurRad="127000" rotWithShape="0">
      <a:srgbClr val="95939D">
        <a:alpha val="16000"/>
      </a:srgbClr>
    </a:outerShdw>
    Rule: If the shape contains opacity (< 100%), it must not have any shadow.
    """
    spPr = shape._element.spPr
    for child in list(spPr):
        if child.tag.endswith("effectLst"):
            spPr.remove(child)

    # Rule: If the shape contains opacity then that shape should not have any shadow.
    alpha_elem = spPr.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}alpha")
    if alpha_elem is not None:
        val = int(alpha_elem.get("val", "100000"))
        if val < 100000:
            return

    if proto_shape is not None:
        p_eff = proto_shape._element.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}effectLst")
        if p_eff is not None:
            spPr.append(copy.deepcopy(p_eff))
            return

    eff_xml = (
        '<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:outerShdw blurRad="127000" rotWithShape="0">'
        '<a:srgbClr val="95939D">'
        '<a:alpha val="16000"/>'
        '</a:srgbClr>'
        '</a:outerShdw>'
        '</a:effectLst>'
    )
    spPr.append(parse_xml(eff_xml))


def load_library():
    """Loads the library deck (falls back to the inspect copy) and returns (lib_prs, lib_s1, lib_s3, lib_lines)."""
    path = LIBRARY_PATH
    if not path.exists() or os.access(path, os.R_OK) is False:
        path = LIBRARY_INSPECT_PATH
    try:
        with open(path, "rb") as f:
            lib_prs = pptx.Presentation(io.BytesIO(f.read()))
    except Exception:
        lib_prs = pptx.Presentation(LIBRARY_INSPECT_PATH)

    # Search dynamically across all slides for the 5 authentic line prototypes
    lib_lines = {}
    for key, name in LIBRARY_LINE_NAMES.items():
        found = None
        for s in lib_prs.slides:
            sh = get_shape_by_name(s, name)
            if sh is not None:
                found = sh
                break
        lib_lines[key] = found

    return lib_prs, lib_prs.slides[0], lib_prs.slides[2], lib_lines


def get_crd_icn_v1_protos(lib_prs):
    """Retrieves prototypes for CRD-ICN-V1 (Data box design with icon circle in header, Library Slide 7)."""
    names = {
        "decor": "CRD-ICN-V1 Decor Box",
        "header": "CRD-ICN-V1 Header Text",
        "icon": "CRD-ICN-V1 Icon Box",
        "divider": "Straight Connector 8",
        "body": "CRD-ICN-V1 Body Text",
    }
    protos = {}
    for s in lib_prs.slides:
        for k, n in names.items():
            if k not in protos:
                sh = get_shape_by_name(s, n)
                if sh is not None:
                    protos[k] = sh
    return protos


def get_imp_box_v1_protos(lib_prs):
    """Retrieves prototypes for IMP-BOX-V1 (Important Box Design with top header tab, Library Slide 8)."""
    names = {
        "decor": "IMP-BOX-V1 Decor Box",
        "header": "IMP-BOX-V1 Header Decor Box",
        "body": "IMP-BOX-V1 Body Text 1",
        "divider": "White Side Fade R Divider",
    }
    protos = {}
    for s in lib_prs.slides:
        for k, n in names.items():
            if k not in protos:
                sh = get_shape_by_name(s, n)
                if sh is not None:
                    protos[k] = sh
    return protos


def set_cell_border(cell, color="E2E8F0", width="6350"):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.set("anchor", "ctr")
    for border_name in ["lnL", "lnR", "lnT", "lnB"]:
        existing = tcPr.find(f"{{http://schemas.openxmlformats.org/drawingml/2006/main}}{border_name}")
        if existing is not None:
            tcPr.remove(existing)
        ln_xml = f'<a:{border_name} xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="{width}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:{border_name}>'
        tcPr.append(parse_xml(ln_xml))


def add_connector_line(slide, x1, y1, x2, y2, color_hex="A2ABB0", width_pt=1.0):
    conn = slide.shapes.add_connector(pptx.enum.shapes.MSO_CONNECTOR.STRAIGHT, Pt(x1), Pt(y1), Pt(x2), Pt(y2))
    conn.line.color.rgb = hex_to_rgb(color_hex)
    conn.line.width = Pt(width_pt)
    return conn


def inject_svg_assets(raw_pptx_bytes, after_copy_path=AFTER_PATH, targets=None):
    """Embeds the 5 SVG icons and adds their relationships to the slides that reference them.
    `targets` maps slide number -> relationship-id prefix used by that slide's builder
    (default: the original deck's slide 2 and slide 4)."""
    targets = targets or {2: "rIdSvgS2_", 4: "rIdSvgS4_"}
    rels_targets = {f"ppt/slides/_rels/slide{n}.xml.rels": p for n, p in targets.items()}
    with zipfile.ZipFile(after_copy_path, "r") as z_in:
        svg_map = {f"image{i}.svg": z_in.read(f"ppt/media/image{i}.svg") for i in range(19, 24)}

    in_buf = io.BytesIO(raw_pptx_bytes)
    out_buf = io.BytesIO()

    with zipfile.ZipFile(in_buf, "r") as zin, zipfile.ZipFile(out_buf, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            # Add SVG relationships to the slides whose builders reference them
            if item.filename in rels_targets:
                prefix = rels_targets[item.filename]
                root = ET.fromstring(data)
                for idx in range(5):
                    rel_id = f"{prefix}{idx+1}"
                    if not any(c.attrib.get("Id") == rel_id for c in root):
                        ET.SubElement(root, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship", {
                            "Id": rel_id,
                            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                            "Target": f"../media/image{19+idx}.svg"
                        })
                data = ET.tostring(root, xml_declaration=True, encoding="utf-8")
            elif item.filename == "[Content_Types].xml":
                root = ET.fromstring(data)
                has_svg = any(c.attrib.get("Extension") == "svg" for c in root)
                if not has_svg:
                    ET.SubElement(root, "{http://schemas.openxmlformats.org/package/2006/content-types}Default", {
                        "Extension": "svg",
                        "ContentType": "image/svg+xml"
                    })
                data = ET.tostring(root, xml_declaration=True, encoding="utf-8")
            zout.writestr(item, data)
        # Add SVGs to media/
        for fn, svg_bytes in svg_map.items():
            zout.writestr(f"ppt/media/{fn}", svg_bytes)

    return out_buf.getvalue()


def clean_bullet_string(txt: str) -> str:
    """Strips literal bullet markers and leading whitespace from text strings."""
    return re.sub(r'^[•\-\*\u2022\u25E6\u25AA\s]+', '', txt).strip()


def apply_native_bullet(paragraph, char="•", typeface="Arial", mar_l_pt=13.5, indent_pt=-13.5):
    """
    Applies PowerPoint's internal native bullet feature to a paragraph via OpenXML DrawingML.
    Strict schema sequence: buFont -> buChar inserted BEFORE tabLst / defRPr / extLst.
    """
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("marL", str(int(mar_l_pt * 12700)))
    pPr.set("indent", str(int(indent_pt * 12700)))

    for child in list(pPr):
        if any(child.tag.endswith(t) for t in (
            "buNone", "buAutoNum", "buChar", "buBlip", "buFont", "buFontTx",
            "buClr", "buClrTx", "buSzPct", "buSzPts", "buSzTx"
        )):
            pPr.remove(child)

    bu_font = parse_xml(
        f'<a:buFont xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" typeface="{typeface}"/>'
    )
    bu_char = parse_xml(
        f'<a:buChar xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" char="{char}"/>'
    )

    insert_idx = len(pPr)
    for i, child in enumerate(pPr):
        if child.tag.endswith(("tabLst", "defRPr", "extLst")):
            insert_idx = i
            break
    pPr.insert(insert_idx, bu_font)
    pPr.insert(insert_idx + 1, bu_char)


def enforce_middle_alignment_deck_wide(prs) -> int:
    """
    Ensures 100% of text boxes and shapes containing text frames across all slides
    have vertical anchor set to Middle (anchor='ctr' / MSO_ANCHOR.MIDDLE).
    """
    count = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                shape.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
                shape.text_frame._txBody.bodyPr.set("anchor", "ctr")
                count += 1
    return count


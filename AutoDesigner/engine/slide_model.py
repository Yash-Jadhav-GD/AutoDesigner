"""
AutoDesigner Slide Model
Reads every shape on a slide (recursing into groups, reading tables cell by cell)
into flat Element records with absolute geometry in points, typography and fills.
Everything downstream (rules mapper, Claude mapper, eval coverage) works from this.
"""

import re
from dataclasses import dataclass, field, asdict
from statistics import median
from typing import Dict, List, Optional

from pptx.enum.dml import MSO_FILL
from pptx.enum.shapes import MSO_SHAPE_TYPE

EMU_PER_PT = 12700
_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
WORD_RE = re.compile(r"\w+", re.UNICODE)


@dataclass
class Para:
    text: str
    level: int = 0
    size: Optional[float] = None   # pt; None when inherited and not resolvable
    bold: bool = False
    color: Optional[str] = None    # hex, or None for theme/inherited


@dataclass
class Element:
    id: int
    kind: str                      # text | table | picture | chart | connector | shape
    name: str
    x: float
    y: float
    w: float
    h: float
    z: int
    placeholder: Optional[str] = None
    paras: List[Para] = field(default_factory=list)
    fill: Optional[str] = None
    line: Optional[str] = None
    table: Optional[List[List[str]]] = None
    group: Optional[int] = None    # id of the enclosing group element, if any

    @property
    def text(self) -> str:
        if self.table is not None:
            return "\n".join(" | ".join(r) for r in self.table)
        return "\n".join(p.text for p in self.paras if p.text)

    @property
    def size(self) -> Optional[float]:
        sizes = [p.size for p in self.paras if p.size and p.text]
        return max(sizes) if sizes else None

    @property
    def bold(self) -> bool:
        texts = [p for p in self.paras if p.text]
        return bool(texts) and all(p.bold for p in texts)

    @property
    def cx(self) -> float:
        return self.x + self.w / 2

    @property
    def cy(self) -> float:
        return self.y + self.h / 2

    def contains(self, other: "Element", tol: float = 2.0) -> bool:
        return (other.x >= self.x - tol and other.y >= self.y - tol
                and other.x + other.w <= self.x + self.w + tol
                and other.y + other.h <= self.y + self.h + tol)


@dataclass
class SlideModel:
    index: int
    width: float
    height: float
    elements: List[Element] = field(default_factory=list)
    notes: str = ""

    @property
    def text_elements(self) -> List[Element]:
        return [e for e in self.elements if e.kind in ("text", "table") and e.text.strip()]

    def reading_order(self, elements: Optional[List[Element]] = None, band: float = 12.0) -> List[Element]:
        """Top-to-bottom row bands, left-to-right within a band."""
        els = sorted(elements if elements is not None else self.text_elements, key=lambda e: e.y)
        rows, cur, top = [], [], None
        for e in els:
            if top is None or e.y - top <= band:
                cur.append(e)
                top = e.y if top is None else top
            else:
                rows.append(cur)
                cur, top = [e], e.y
        if cur:
            rows.append(cur)
        return [e for r in rows for e in sorted(r, key=lambda e: e.x)]

    def median_size(self) -> Optional[float]:
        sizes = [e.size for e in self.text_elements if e.size]
        return median(sizes) if sizes else None

    def title(self) -> Optional[Element]:
        """Title placeholder first; otherwise the largest-font text in the top 15% of the slide
        whose (size, bold) style does not repeat down the same column (that is a subheader list)."""
        for e in self.text_elements:
            if e.placeholder in ("TITLE", "CENTER_TITLE"):
                return e
        texts = [e for e in self.text_elements if e.kind == "text"]
        top = [e for e in texts if e.y < self.height * 0.15]
        if not top:
            return None

        def style(e):
            return (e.size, e.bold)

        unique = [e for e in top if e.size
                  and not any(o is not e and style(o) == style(e) and abs(o.x - e.x) < 5 for o in texts)]
        med = self.median_size()
        if unique:
            best = max(unique, key=lambda e: (e.size or 0, e.w))
            if med and best.size > med:
                return best
        # Without size information, accept a single wide text in the top band.
        wide = [e for e in top if e.w > self.width * 0.5 and not e.size]
        return wide[0] if len(wide) == 1 else None

    def words(self) -> List[str]:
        out = []
        for e in self.text_elements:
            out.extend(w.lower() for w in WORD_RE.findall(e.text))
        return out

    def to_dict(self) -> Dict:
        return asdict(self)


# ------------------------------------------------------------------------------
# Reading
# ------------------------------------------------------------------------------

def _pt(v) -> float:
    return round((v or 0) / EMU_PER_PT, 1)


def _solid_hex(fmt) -> Optional[str]:
    try:
        if fmt.type == MSO_FILL.SOLID:
            return str(fmt.fore_color.rgb)
    except Exception:
        pass
    return None


def _line_hex(shape) -> Optional[str]:
    try:
        return str(shape.line.color.rgb) if shape.line.fill.type == MSO_FILL.SOLID else None
    except Exception:
        return None


def _paras(text_frame) -> List[Para]:
    out = []
    for p in text_frame.paragraphs:
        text = p.text.strip()
        size = bold = color = None
        for r in p.runs:
            if not r.text.strip():
                continue
            if size is None and r.font.size is not None:
                size = r.font.size.pt
            if bold is None and r.font.bold is not None:
                bold = r.font.bold
            if color is None:
                try:
                    color = str(r.font.color.rgb) if r.font.color and r.font.color.type is not None else None
                except Exception:
                    color = None
        if size is None and p.font.size is not None:
            size = p.font.size.pt
        if bold is None and p.font.bold is not None:
            bold = p.font.bold
        if size is None:
            # endParaRPr / defRPr sometimes carry the size when runs do not
            for tag in ("endParaRPr", "pPr/" + _A + "defRPr"):
                el = p._p.find(_A + tag)
                if el is not None and el.get("sz"):
                    size = int(el.get("sz")) / 100
                    break
        out.append(Para(text=text, level=p.level, size=size, bold=bool(bold), color=color))
    return out


def _group_transform(group_shape):
    """Returns f(x, y, w, h) mapping child EMU coordinates into slide EMU coordinates."""
    xfrm = group_shape._element.grpSpPr.find(_A + "xfrm")
    if xfrm is None:
        return lambda x, y, w, h: (x, y, w, h)
    off, ext = xfrm.find(_A + "off"), xfrm.find(_A + "ext")
    choff, chext = xfrm.find(_A + "chOff"), xfrm.find(_A + "chExt")
    if None in (off, ext, choff, chext):
        return lambda x, y, w, h: (x, y, w, h)
    ox, oy = int(off.get("x")), int(off.get("y"))
    cx, cy = int(ext.get("cx")), int(ext.get("cy"))
    chx, chy = int(choff.get("x")), int(choff.get("y"))
    ccx, ccy = int(chext.get("cx")) or 1, int(chext.get("cy")) or 1
    sx, sy = cx / ccx, cy / ccy
    return lambda x, y, w, h: (ox + (x - chx) * sx, oy + (y - chy) * sy, w * sx, h * sy)


def read_slide(slide, index: int, slide_width: int, slide_height: int) -> SlideModel:
    model = SlideModel(index=index, width=_pt(slide_width), height=_pt(slide_height))
    counter = [0]

    def walk(shapes, transform, parent_group):
        for shape in shapes:
            counter[0] += 1
            eid = counter[0]
            x, y, w, h = transform(shape.left or 0, shape.top or 0, shape.width or 0, shape.height or 0)
            base = dict(id=eid, name=shape.name, x=_pt(x), y=_pt(y), w=_pt(w), h=_pt(h), z=eid, group=parent_group)

            st = shape.shape_type
            if st == MSO_SHAPE_TYPE.GROUP:
                model.elements.append(Element(kind="group", **base))
                child_t = _group_transform(shape)
                walk(shape.shapes, lambda a, b, c, d, ct=child_t, pt=transform: pt(*ct(a, b, c, d)), eid)
                continue

            placeholder = None
            if shape.is_placeholder:
                try:
                    placeholder = shape.placeholder_format.type.name if shape.placeholder_format.type else None
                except Exception:
                    placeholder = None

            if getattr(shape, "has_table", False) and shape.has_table:
                grid = []
                for r in shape.table.rows:
                    grid.append(["\n".join(p.text.strip() for p in c.text_frame.paragraphs if p.text.strip())
                                 for c in r.cells])
                model.elements.append(Element(kind="table", table=grid, placeholder=placeholder, **base))
                continue
            if getattr(shape, "has_chart", False) and shape.has_chart:
                model.elements.append(Element(kind="chart", placeholder=placeholder, **base))
                continue
            if st == MSO_SHAPE_TYPE.PICTURE:
                model.elements.append(Element(kind="picture", placeholder=placeholder, **base))
                continue
            if st in (MSO_SHAPE_TYPE.LINE,) or shape._element.tag.endswith("cxnSp"):
                model.elements.append(Element(kind="connector", line=_line_hex(shape), **base))
                continue

            fill = _solid_hex(shape.fill) if hasattr(shape, "fill") else None
            paras = _paras(shape.text_frame) if getattr(shape, "has_text_frame", False) and shape.has_text_frame else []
            kind = "text" if any(p.text for p in paras) else "shape"
            model.elements.append(Element(kind=kind, paras=paras, fill=fill, line=_line_hex(shape),
                                          placeholder=placeholder, **base))

    walk(slide.shapes, lambda x, y, w, h: (x, y, w, h), None)

    try:
        if slide.has_notes_slide:
            model.notes = slide.notes_slide.notes_text_frame.text.strip()
    except Exception:
        pass
    return model


def read_presentation(prs) -> List[SlideModel]:
    return [read_slide(s, i + 1, prs.slide_width, prs.slide_height) for i, s in enumerate(prs.slides)]


# ------------------------------------------------------------------------------
# Role clustering primitives (AUTOFORMAT_SPECIFICATION.md section 4)
# ------------------------------------------------------------------------------

def detect_background(elements: List[Element]) -> Optional[Element]:
    """Spec 4.2: the largest shape is a background container if wider than 1.5x the average width."""
    if len(elements) < 2:
        return None
    biggest = max(elements, key=lambda e: e.w * e.h)
    avg_w = sum(e.w for e in elements) / len(elements)
    return biggest if biggest.w > 1.5 * avg_w else None


def rows_by_midpoint(elements: List[Element]) -> List[List[Element]]:
    """Spec 4.4: a shape joins the current row when its mid-Y falls inside the previous shape's span."""
    rows: List[List[Element]] = []
    for e in sorted(elements, key=lambda e: e.y):
        prev = rows[-1][-1] if rows else None
        if prev is not None and prev.y <= e.cy <= prev.y + prev.h:
            rows[-1].append(e)
        else:
            rows.append([e])
    return [sorted(r, key=lambda e: e.x) for r in rows]


def pair_sequential(elements: List[Element]) -> List[tuple]:
    """Spec 4.6: sorted by top, items 2k and 2k+1 form (subheader, body)."""
    els = sorted(elements, key=lambda e: e.y)
    return [(els[i], els[i + 1] if i + 1 < len(els) else None) for i in range(0, len(els), 2)]


def columns_by_gap(elements: List[Element], min_gap: float = 15.0) -> List[List[Element]]:
    """Clusters elements into columns where their horizontal extents are separated by >= min_gap."""
    cols: List[List[Element]] = []
    right = None
    for e in sorted(elements, key=lambda e: e.x):
        if right is not None and e.x - right < min_gap:
            cols[-1].append(e)
            right = max(right, e.x + e.w)
        else:
            cols.append([e])
            right = e.x + e.w
    return [sorted(c, key=lambda e: e.y) for c in cols]


def split_by_largest_gap(elements: List[Element], min_gap: float = 20.0):
    """Splits into (left, right) at the widest vertical gutter, or returns None when there is none."""
    cols = columns_by_gap(elements, min_gap)
    if len(cols) < 2:
        return None
    gaps = []
    for i in range(1, len(cols)):
        left_edge = max(e.x + e.w for c in cols[:i] for e in c)
        right_edge = min(e.x for e in cols[i])
        gaps.append((right_edge - left_edge, i))
    _, cut = max(gaps)
    return [e for c in cols[:cut] for e in c], [e for c in cols[cut:] for e in c]


# ------------------------------------------------------------------------------
# Coverage
# ------------------------------------------------------------------------------

def spec_words(data) -> List[str]:
    """All words in any string inside a spec data structure."""
    out = []
    if isinstance(data, str):
        out.extend(w.lower() for w in WORD_RE.findall(data))
    elif isinstance(data, dict):
        for v in data.values():
            out.extend(spec_words(v))
    elif isinstance(data, (list, tuple)):
        for v in data:
            out.extend(spec_words(v))
    return out


def coverage(model: SlideModel, data) -> Dict:
    """Fraction of source words (multiset) that appear in the spec data, plus the missing ones."""
    from collections import Counter
    src, got = Counter(model.words()), Counter(spec_words(data))
    missing = src - got
    total = sum(src.values())
    covered = total - sum(missing.values())
    return {
        "coverage": 1.0 if total == 0 else covered / total,
        "missing": sorted(missing.elements()),
    }

"""
AutoDesigner Layout Engine
Geometry for the generic card archetypes. Every card offset and size is read from the
Library_copy.pptx prototype shapes (found by name) and combined with the layout formulas in
AUTOFORMAT_SPECIFICATION.md section 5: header offsets and protrusion, 0.3 cm padding,
divider gap distribution and aspect-ratio-locked expansion.
"""

import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

try:
    from .analyzer import SlideData, CardData
    from .color_tokens import TypographyTokens as T
except ImportError:
    from engine.analyzer import SlideData, CardData
    from engine.color_tokens import TypographyTokens as T


# Styler archetype -> role -> Library prototype name (AUTOFORMAT_SPECIFICATION.md section 3).
# A design's prototypes are looked up on the Library slide that holds its decor box.
DESIGNS: Dict[str, Dict[str, str]] = {
    "design1_floating_pill": {          # Library slide 1
        "decor": "Design 1 Decor box",
        "header": "Design 1 Header Text",
        "body": "Design 1 Body text 1",
    },
    "icon_badge_card": {                # Library slide 4
        "decor": "Design 1 Decor",
        "header": "Design 1 Header Text",
        "body": "Design 1 Body Text 1",
        "badge": "Tiny CircleBackplate 25",
    },
    "crd_sub_lst": {                    # Library slide 2
        "decor": "Design 1 CRD-SUB-LST Decor box",
        "header": "Design 1 CRD-SUB-LST Header Text",
        "sub": "Design 1 CRD-SUB-LST Sub header Text 1",
        "body": "Design 1 CRD-SUB-LST Body text 1",
    },
    "kpi_metric_card": {                # Library slide 5
        "decor": "KPI-CRD-V1 Decor box",
        "number": "KPI-CRD-V1 Number",
        "header": "KPI-CRD-V1 Header Text",
        "body": "KPI-CRD-V1 Body Text",
    },
    "crd_icn_v1": {                     # Library slide 7: Data box design
        "decor": "CRD-ICN-V1 Decor Box",
        "header": "CRD-ICN-V1 Header Text",
        "icon": "CRD-ICN-V1 Icon Box",
        "body": "CRD-ICN-V1 Body Text",
        "divider": "Straight Connector 8",
    },
    "imp_box_v1": {                     # Library slide 8: Important Box Design
        "decor": "IMP-BOX-V1 Decor Box",
        "header": "IMP-BOX-V1 Header Decor Box",
        "body": "IMP-BOX-V1 Body Text 1",
        "divider": "White Side Fade R Divider",
    },
}
# Subheader + body pair, used for items with a subheader in any list design
SUBHEADER_PROTO = "Design 1 CRD-SUB-LST Sub header Text 1"
SUB_BODY_PROTO = "Design 1 CRD-SUB-LST Body text 1"
DIVIDER_PROTO = "Grey Side Fade R"      # Library slide 6


@dataclass
class Rect:
    left: int
    top: int
    width: int
    height: int


def _rect(left_pt: float, top_pt: float, width_pt: float, height_pt: float) -> Rect:
    return Rect(left=Pt(left_pt), top=Pt(top_pt), width=Pt(width_pt), height=Pt(height_pt))


@dataclass
class Proto:
    """A Library prototype shape with its geometry and text insets in pt."""
    name: str
    shape: object
    left: float
    top: float
    width: float
    height: float
    insets: Tuple[float, float, float, float]  # left, top, right, bottom

    @property
    def right(self) -> float:
        return self.left + self.width

    @property
    def bottom(self) -> float:
        return self.top + self.height


def _norm(name: str) -> str:
    return re.sub(r"\s+", "", name).lower()


def _insets(shape) -> Tuple[float, float, float, float]:
    defaults = (("lIns", 7.2), ("tIns", 3.6), ("rIns", 7.2), ("bIns", 3.6))
    body_pr = shape._element.find(".//" + qn("a:bodyPr"))
    return tuple(Emu(int(body_pr.get(k))).pt if body_pr is not None and body_pr.get(k) else v
                 for k, v in defaults)


class Prototypes:
    """Library prototype shapes by name (whitespace and case are ignored)."""

    def __init__(self, lib_prs):
        self.slides = [{_norm(s.name): s for s in slide.shapes} for slide in lib_prs.slides]

    def proto(self, name: str, slide: Optional[dict] = None) -> Proto:
        key = _norm(name)
        shape = (slide or {}).get(key) or next((s[key] for s in self.slides if key in s), None)
        if shape is None:
            raise KeyError(f"Library prototype {name!r} not found in Library_copy.pptx")
        return Proto(name, shape, Emu(shape.left).pt, Emu(shape.top).pt,
                     Emu(shape.width).pt, Emu(shape.height).pt, _insets(shape))

    def design(self, archetype: str) -> Dict[str, Proto]:
        names = DESIGNS[archetype]
        anchor = _norm(names["decor"])
        slide = next((s for s in self.slides if anchor in s), None)
        if slide is None:
            raise KeyError(f"Library prototype {names['decor']!r} not found in Library_copy.pptx")
        design = {role: self.proto(name, slide) for role, name in names.items()}
        if archetype != "kpi_metric_card":
            design.setdefault("sub", self.proto(SUBHEADER_PROTO))
            design.setdefault("sub_body", self.proto(SUB_BODY_PROTO))
            if "divider" not in names:
                design["divider"] = self.proto(DIVIDER_PROTO)
        return design


def text_lines(text: str, width_pt: float, size_pt: float, bold: bool = False) -> int:
    """Greedy word-wrap estimate of the number of lines `text` takes in `width_pt`."""
    em = T.CHAR_WIDTH_EM_BOLD if bold else T.CHAR_WIDTH_EM
    cap = max(1, int(width_pt / (size_pt * em)))
    total = 0
    for para in text.split("\n"):
        lines, cur = 1, 0
        for word in para.split():
            n = len(word)
            if cur and cur + 1 + n <= cap:
                cur += 1 + n
                continue
            if cur:
                lines += 1
            lines += (n - 1) // cap
            cur = n - ((n - 1) // cap) * cap
        total += lines
    return total


def text_height(proto: Proto, paragraphs: List[str], width_pt: float, size_pt: float,
                bold: bool = False, min_h: float = 0.0) -> float:
    left, top, right, bottom = proto.insets
    lines = sum(text_lines(p, width_pt - left - right, size_pt, bold) for p in paragraphs)
    return max(min_h, lines * size_pt * T.LINE_SPACING + top + bottom)


@dataclass
class TextBox:
    rect: Rect
    paragraphs: List[str]
    size: float


@dataclass
class RowLayout:
    index: int
    top: int
    height: int
    text: Optional[TextBox]                     # body text
    sub: Optional[TextBox] = None               # subheader above the body
    badge_rect: Optional[Rect] = None           # icon backplate
    icon_rect: Optional[Rect] = None            # SVG icon centred in the backplate
    icon_text: str = ""                         # text the icon is matched against
    divider: Optional[Tuple[float, float, float]] = None  # left, top, width in pt (after this row)


@dataclass
class CardLayout:
    card_index: int
    container_rect: Rect
    header: Optional[TextBox]
    header_style: str  # "floating_pill", "flush", "kpi", "header_icon", "top_tab"
    rows: List[RowLayout]
    number: Optional[TextBox] = None
    header_icon_rect: Optional[Rect] = None
    header_icon_text: str = ""
    header_divider: Optional[Tuple[float, float, float]] = None


@dataclass
class SlideLayout:
    slide_width: int
    slide_height: int
    title_rect: Rect
    subtitle_rect: Rect
    category_rect: Rect
    cards: List[CardLayout]
    design: Dict[str, Proto] = None


class LayoutEngine:
    """Calculates spatial coordinates for 16:9 slides (960pt x 540pt)."""

    SLIDE_WIDTH_PT = 960
    SLIDE_HEIGHT_PT = 540

    CANVAS_LEFT_PT = 45
    CANVAS_RIGHT_PT = 915
    USABLE_WIDTH_PT = 870  # 915 - 45

    TOP_TITLE_PT = 26
    TITLE_HEIGHT_PT = 46

    CONTENT_TOP_PT = 96            # top of the card cluster (a protruding header starts here)
    CONTENT_TOP_NO_TITLE_PT = 40
    CONTENT_BOTTOM_PT = 512

    MAX_COLUMNS = 4
    COLUMN_GAPS_PT = {2: 36.0, 3: 26.0, 4: 18.0}
    SINGLE_CARD_WIDTH_PT = 580.0
    ROW_GAP_PT = 24.0
    MAX_ROW_COLUMNS = 3           # dense cards flow their rows into up to 3 columns
    MIN_ROW_COLUMN_PT = 170.0     # narrowest row column

    @classmethod
    def compute_slide_layout(cls, slide_data: SlideData, archetype: str, prototypes: Prototypes,
                             has_category: bool = False) -> SlideLayout:
        design = prototypes.design(archetype)
        has_title = bool(slide_data.title.strip())
        title_rect, cat_rect, sub_rect = cls._title_rects(has_title, has_category, bool(slide_data.subtitle))

        if not slide_data.cards:
            return SlideLayout(Pt(cls.SLIDE_WIDTH_PT), Pt(cls.SLIDE_HEIGHT_PT), title_rect, sub_rect, cat_rect, [], design)

        top = cls.CONTENT_TOP_PT if has_title else cls.CONTENT_TOP_NO_TITLE_PT
        slot_rows = cls._slot_rows(len(slide_data.cards))
        band_h = (cls.CONTENT_BOTTOM_PT - top - (len(slot_rows) - 1) * cls.ROW_GAP_PT) / max(1, len(slot_rows))
        is_kpi = archetype == "kpi_metric_card"
        place = cls._kpi_card if is_kpi else cls._list_card

        # Fit: at the largest body size possible, each card takes the first placement option
        # (slot and number of row columns) that fits its band. Text is never dropped, so the
        # smallest size and widest option are used even if the content still overflows.
        sizes = [T.BODY - 0.5 * i for i in range(int((T.BODY - T.BODY_MIN) / 0.5) + 1)]
        for size in sizes:
            plan, fits = [], True
            cards = iter(slide_data.cards)
            for slots in slot_rows:
                row = []
                for slot in slots:
                    card = next(cards)
                    for x, w, ncols in cls._placement_options(slot, len(slide_data.cards), is_kpi, design):
                        m = place(card, x, top, w, band_h, design, size, ncols=ncols)
                        if m[0] <= m[2]:
                            break
                    else:
                        fits = False
                    row.append((x, w, ncols, m))
                plan.append(row)
            if fits:
                break

        cards: List[CardLayout] = []
        card_iter = iter(enumerate(slide_data.cards))
        y = top
        for row in plan:
            # Uniform height across a row of cards: template-ratio height, grown to fit content
            bg_h = max(max(req, base) for *_, (req, base, _) in row)
            for x, w, ncols, _ in row:
                idx, card = next(card_iter)
                cards.append(place(card, x, y, w, band_h, design, size, bg_h=bg_h, idx=idx, ncols=ncols))
            row_bottom = max(Emu(c.container_rect.top + c.container_rect.height).pt for c in cards[-len(row):])
            y = row_bottom + cls.ROW_GAP_PT

        return SlideLayout(
            slide_width=Pt(cls.SLIDE_WIDTH_PT),
            slide_height=Pt(cls.SLIDE_HEIGHT_PT),
            title_rect=title_rect,
            subtitle_rect=sub_rect,
            category_rect=cat_rect,
            cards=cards,
            design=design,
        )

    @classmethod
    def _title_rects(cls, has_title: bool, has_category: bool, has_subtitle: bool):
        if not has_title:
            return None, None, None
        title_y = cls.TOP_TITLE_PT
        cat_rect = None
        if has_category:
            cat_rect = _rect(cls.CANVAS_LEFT_PT, 14, 200, 18)
            title_y = 34
        title_rect = _rect(cls.CANVAS_LEFT_PT, title_y, cls.USABLE_WIDTH_PT, cls.TITLE_HEIGHT_PT)
        sub_rect = _rect(cls.CANVAS_LEFT_PT, title_y + 36, cls.USABLE_WIDTH_PT, 22) if has_subtitle else None
        return title_rect, cat_rect, sub_rect

    @classmethod
    def _list_card(cls, card: CardData, x: float, y: float, slot_w: float, max_h: float,
                   d: Dict[str, Proto], size: float, bg_h: Optional[float] = None, idx: int = 0,
                   ncols: int = 1):
        """Header + rows card (Library slides 1, 2, 4, 7 and 8), rows flowing into `ncols` columns.
        Without `bg_h` returns (required height, template-ratio height, max height) for the
        background container."""
        decor, header, body, badge, divider = d["decor"], d["header"], d["body"], d.get("badge"), d["divider"]
        icon_proto = d.get("icon")
        is_header_icon = icon_proto is not None
        is_top_tab = header.name == "IMP-BOX-V1 Header Decor Box"
        pad = T.CARD_PADDING

        ratio = decor.height / decor.width
        bg_w = slot_w

        header_icon_rect = None
        header_icon_text = ""
        header_divider = None
        header_box = None

        if is_header_icon:
            # CRD-ICN-V1: Icon circle in header, divider line below header, body text below divider
            header_style = "header_icon"
            has_row_divs = False
            bg_off_top = 0.0
            bg_top = y

            icon_w = icon_proto.width
            icon_off_l = icon_proto.left - decor.left
            icon_off_t = icon_proto.top - decor.top
            header_icon_rect = _rect(x + icon_off_l, bg_top + icon_off_t, icon_w, icon_w)
            header_icon_text = card.header

            hw = bg_w - (icon_off_l + icon_w + 11.0) - 12.0
            hh = text_height(header, [card.header], hw, T.KPI_HEADER, bold=True, min_h=22.0)
            hy = bg_top + icon_off_t + (icon_w - hh) / 2
            hx = x + icon_off_l + icon_w + 11.0
            header_box = TextBox(_rect(hx, hy, hw, hh), [card.header], T.KPI_HEADER)

            div_off_l = divider.left - decor.left
            div_y = bg_top + icon_off_t + icon_w + 8.5
            div_w = bg_w - 2 * div_off_l
            header_divider = (x + div_off_l, div_y, div_w)
            start_y = div_y + 9.0

            body_off_l = body.left - decor.left if body else 10.0
            body_off_r = decor.right - body.right if body else 10.0
            line_off_l = body_off_l

        elif is_top_tab:
            # IMP-BOX-V1: round2SameRect top tab centered horizontally, white fading dividers between body rows
            header_style = "top_tab"
            has_row_divs = True
            bg_off_top = 0.0
            bg_top = y

            tab_w = min(210.0, bg_w * 0.65)
            tab_h = header.height
            tab_x = x + (bg_w - tab_w) / 2
            header_box = TextBox(_rect(tab_x, bg_top, tab_w, tab_h), [card.header], 14.0)
            start_y = bg_top + tab_h + 10.0

            body_off_l = 18.0
            body_off_r = 18.0
            line_off_l = body_off_l

        else:
            # Standard floating pill or flush header
            header_off_l = header.left - decor.left
            header_off_r = decor.right - header.right
            bg_off_top = decor.top - header.top
            body_off_l = body.left - decor.left
            body_off_r = decor.right - body.right
            badge_off_l = badge.left - decor.left if badge else body_off_l
            line_off_l = min(body_off_l, badge_off_l)
            has_row_divs = True
            header_style = "floating_pill" if bg_off_top > 0 else "flush"

            if bg_off_top > 0:   # protruding header (Design 1)
                bg_top, header_top = y + bg_off_top, y
            else:                # flush header (Design 2 / CRD-SUB-LST)
                bg_top, header_top = y, y - bg_off_top

            start_y = bg_top + pad
            if card.header.strip():
                hw = bg_w - header_off_l - header_off_r
                hh = text_height(header, [card.header], hw, T.PILL_HEADER, bold=True, min_h=header.height)
                header_box = TextBox(_rect(x + header_off_l, header_top, hw, hh), [card.header], T.PILL_HEADER)
                start_y = header_top + hh + pad

        max_bg_h = max_h - (bg_top - y)
        base_h = min(max(50.0, bg_w * ratio), max_bg_h)

        badge_off_l = badge.left - decor.left if badge else body_off_l
        col_gap = body_off_l + body_off_r
        col_w = (bg_w - line_off_l - body_off_r - (ncols - 1) * col_gap) / ncols
        text_dx = body_off_l - line_off_l
        text_w = col_w - text_dx

        # Pass 1: row heights
        sub_gap = max(0.0, d["sub_body"].top - d["sub"].top - T.SUBHEADER * T.LINE_SPACING)
        rows = []
        for item in card.items:
            if not (item.subheader or item.body):
                continue
            sub_h = text_height(d["sub"], [item.subheader], text_w, T.SUBHEADER, bold=True) if item.subheader else 0.0
            body_h = text_height(body, [item.body], text_w, size) if item.body else 0.0
            gap = sub_gap if sub_h and body_h else 0.0
            text_h = sub_h + gap + body_h
            row_h = max(text_h, badge.height) if badge else text_h
            rows.append((item, sub_h, gap, body_h, text_h, row_h))

        # For line prototypes that are horizontal or 90-deg rotated in XML, thickness is ~1.0 pt
        div_thick = min(divider.width, divider.height) if min(divider.width, divider.height) > 0 else 1.0
        sep_h = div_thick if has_row_divs else 0.0
        columns = cls._split_columns(rows, ncols)

        def stack(col):
            sep_cnt = max(0, len(col) - 1) if has_row_divs else 0
            gap_cnt = max(0, len(col) - 1)
            return sum(r[5] for r in col) + sep_cnt * sep_h + gap_cnt * (2 * T.MIN_GAP)

        required = (start_y - bg_top) + max((stack(c) for c in columns), default=0.0) + pad
        if bg_h is None:
            return required, base_h, max_bg_h

        icon_size = 0.0
        if badge and rows:
            icon_size = max(T.ICON_MIN, min(r[4] for r in rows))
            icon_size = min(icon_size, badge.width * 0.6)

        # Pass 2: vertical space and divider distribution, per row column
        end_y = bg_top + bg_h - pad
        out_rows = []
        bottom = start_y
        for c, col in enumerate(columns):
            col_x = x + line_off_l + c * (col_w + col_gap)
            n = len(col)
            sep_cnt = (n - 1) if has_row_divs else 0
            empty = (end_y - start_y) - sum(r[5] for r in col) - sep_cnt * sep_h
            gap_h = max(T.MIN_GAP, empty / (2 * (n - 1))) if n > 1 else 0.0
            cur = start_y
            for i, (item, sub_h, gap, body_h, text_h, row_h) in enumerate(col):
                text_top = cur + (row_h - text_h) / 2
                text_l = col_x + text_dx
                sub_box = TextBox(_rect(text_l, text_top, text_w, sub_h), [item.subheader], T.SUBHEADER) if sub_h else None
                body_box = TextBox(_rect(text_l, text_top + sub_h + gap, text_w, body_h), [item.body], size) if body_h else None
                badge_rect = icon_rect = None
                if badge:
                    b_top = cur + (row_h - badge.height) / 2
                    b_left = col_x + badge_off_l - line_off_l
                    badge_rect = _rect(b_left, b_top, badge.width, badge.height)
                    icon_rect = _rect(b_left + (badge.width - icon_size) / 2, b_top + (badge.height - icon_size) / 2,
                                      icon_size, icon_size)
                row = RowLayout(index=len(out_rows), top=Pt(cur), height=Pt(row_h), text=body_box, sub=sub_box,
                                badge_rect=badge_rect, icon_rect=icon_rect,
                                icon_text=" ".join(t for t in (item.subheader, item.body) if t))
                cur += row_h
                if i < n - 1:
                    cur += gap_h
                    if has_row_divs:
                        row.divider = (col_x, cur, col_w)
                        cur += sep_h + gap_h
                    else:
                        cur += gap_h
                out_rows.append(row)
            bottom = max(bottom, cur)

        # Pass 3: aspect-ratio-preserving expansion (width grows, bounded by the column slot)
        required_h = bottom + pad - bg_top
        if required_h > bg_h + 0.01:
            bg_h = required_h
            bg_w = min(slot_w, max(bg_w, bg_h / ratio))

        return CardLayout(
            card_index=idx,
            container_rect=_rect(x, bg_top, bg_w, bg_h),
            header=header_box,
            header_style=header_style,
            rows=out_rows,
            header_icon_rect=header_icon_rect,
            header_icon_text=header_icon_text,
            header_divider=header_divider,
        )

    @staticmethod
    def _split_columns(rows: list, ncols: int) -> List[list]:
        """Splits rows, in reading order, into at most `ncols` columns of similar height."""
        if ncols <= 1 or len(rows) <= 1:
            return [rows]
        target = sum(r[5] for r in rows) / ncols
        columns, cur, acc = [], [], 0.0
        for i, r in enumerate(rows):
            remaining = len(rows) - i
            if cur and len(columns) < ncols - 1 and (acc + r[5] / 2 > target or remaining < ncols - len(columns)):
                columns.append(cur)
                cur, acc = [], 0.0
            cur.append(r)
            acc += r[5]
        columns.append(cur)
        return columns

    @classmethod
    def _kpi_card(cls, card: CardData, x: float, y: float, slot_w: float, max_h: float,
                  d: Dict[str, Proto], size: float, bg_h: Optional[float] = None, idx: int = 0,
                  ncols: int = 1):
        """KPI-CRD-V1 card (Library slide 5): number, header label, body text. Section 4 role
        remap: the metric is the number, the card header the label, the rest the body."""
        decor, number, header, body = d["decor"], d["number"], d["header"], d["body"]
        aspect = decor.width / decor.height
        body_size = T.KPI_BODY - (T.BODY - size)

        bg_w = slot_w
        bg_top = y
        base_h = min(max(50.0, bg_w / aspect), max_h)

        def column(p: Proto):
            return x + (p.left - decor.left), bg_w - (p.left - decor.left) - (decor.right - p.right)

        # (prototype, paragraphs, size, bold, gap to the next prototype's top)
        parts = []
        if card.kpi_metric:
            parts.append((number, [card.kpi_metric], T.KPI_NUMBER, True, header.top - number.bottom))
        if card.header.strip():
            parts.append((header, [card.header], T.KPI_HEADER, True, body.top - header.bottom))
        paras = [p for p in [card.kpi_label or ""] + [t for i in card.items for t in (i.subheader, i.body)] if p]
        if paras:
            parts.append((body, paras, body_size, False, 0.0))

        heights = []
        for p, texts, sz, bold, _ in parts:
            _, w = column(p)
            heights.append(text_height(p, texts, w, sz, bold, min_h=p.height if len(texts) == 1 else 0.0))
        lead = number.top - decor.top
        content_h = lead + sum(heights) + sum(g for *_, g in parts[:-1])
        required = content_h + T.CARD_PADDING
        if bg_h is None:
            return required, base_h, max_h

        # Pass 3: aspect-ratio-preserving expansion (width bounded by the column slot)
        if required > bg_h:
            bg_h = required
            bg_w = min(slot_w, bg_h * aspect)

        # Centre the content block vertically, keeping the prototype's bottom margin
        tail = decor.bottom - body.bottom
        cur = bg_top + lead + max(0.0, (bg_h - content_h - tail) / 2)
        boxes = []
        for (p, texts, sz, _, gap), h in zip(parts, heights):
            left, w = column(p)
            boxes.append((p, TextBox(_rect(left, cur, w, h), texts, sz)))
            cur += h + gap

        number_box = next((b for p, b in boxes if p is number), None)
        header_box = next((b for p, b in boxes if p is header), None)
        body_box = next((b for p, b in boxes if p is body), None)
        rows = [RowLayout(index=0, top=body_box.rect.top, height=body_box.rect.height, text=body_box)] if body_box else []
        return CardLayout(
            card_index=idx,
            container_rect=_rect(x, bg_top, bg_w, bg_h),
            header=header_box,
            header_style="kpi",
            rows=rows,
            number=number_box,
        )

    @classmethod
    def _placement_options(cls, slot: Tuple[float, float], count: int, is_kpi: bool,
                           d: Dict[str, Proto]) -> List[Tuple[float, float, int]]:
        """(left, width, row columns) to try for a card, in order of preference: its slot with one
        column, then more row columns; a lone card widens to the full canvas before splitting."""
        x, w = slot
        if is_kpi:
            return [(x, w, 1)]
        options = [(x, w, 1)]
        if count == 1:
            x, w = cls.CANVAS_LEFT_PT, cls.USABLE_WIDTH_PT
            options.append((x, w, 1))
        inset = (d["body"].left - d["decor"].left) + (d["decor"].right - d["body"].right)
        max_cols = min(cls.MAX_ROW_COLUMNS, max(1, int((w - inset) // (cls.MIN_ROW_COLUMN_PT + inset))))
        options += [(x, w, n) for n in range(2, max_cols + 1)]
        return options

    @classmethod
    def _slot_rows(cls, count: int) -> List[List[Tuple[float, float]]]:
        """Column slots (left, width) per row of cards: up to 4 columns, then wrap."""
        count = max(1, count)
        cols = min(count, cls.MAX_COLUMNS)
        if cols == 1:
            w = cls.SINGLE_CARD_WIDTH_PT
            return [[((cls.SLIDE_WIDTH_PT - w) / 2.0, w)]]
        gap = cls.COLUMN_GAPS_PT[cols]
        w = (cls.USABLE_WIDTH_PT - gap * (cols - 1)) / cols
        rows = []
        for start in range(0, count, cols):
            k = min(cols, count - start)
            left = cls.CANVAS_LEFT_PT + (cls.USABLE_WIDTH_PT - (k * w + (k - 1) * gap)) / 2.0
            rows.append([(left + i * (w + gap), w) for i in range(k)])
        return rows

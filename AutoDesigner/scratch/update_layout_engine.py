import re
from pathlib import Path

path = Path("AutoDesigner/engine/layout_engine.py")
text = path.read_text(encoding="utf-8")

# 1. Update DESIGNS
old_designs_end = """    "kpi_metric_card": {                # Library slide 5
        "decor": "KPI-CRD-V1 Decor box",
        "number": "KPI-CRD-V1 Number",
        "header": "KPI-CRD-V1 Header Text",
        "body": "KPI-CRD-V1 Body Text",
    },
}"""

new_designs_end = """    "kpi_metric_card": {                # Library slide 5
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
}"""

assert old_designs_end in text, "old_designs_end not found!"
text = text.replace(old_designs_end, new_designs_end, 1)

# 2. Update Prototypes.design to keep custom divider if defined in design
old_proto_design = """        if archetype != "kpi_metric_card":
            design.setdefault("sub", self.proto(SUBHEADER_PROTO))
            design["sub_body"] = self.proto(SUB_BODY_PROTO)
            design["divider"] = self.proto(DIVIDER_PROTO)
        return design"""

new_proto_design = """        if archetype != "kpi_metric_card":
            design.setdefault("sub", self.proto(SUBHEADER_PROTO))
            design.setdefault("sub_body", self.proto(SUB_BODY_PROTO))
            if "divider" not in names:
                design["divider"] = self.proto(DIVIDER_PROTO)
        return design"""

assert old_proto_design in text, "old_proto_design not found!"
text = text.replace(old_proto_design, new_proto_design, 1)

# 3. Update CardLayout
old_card_layout = """@dataclass
class CardLayout:
    card_index: int
    container_rect: Rect
    header: Optional[TextBox]
    header_style: str  # "floating_pill", "flush", "kpi"
    rows: List[RowLayout]
    number: Optional[TextBox] = None"""

new_card_layout = """@dataclass
class CardLayout:
    card_index: int
    container_rect: Rect
    header: Optional[TextBox]
    header_style: str  # "floating_pill", "flush", "kpi", "header_icon", "top_tab"
    rows: List[RowLayout]
    number: Optional[TextBox] = None
    header_icon_rect: Optional[Rect] = None
    header_icon_text: str = ""
    header_divider: Optional[Tuple[float, float, float]] = None"""

assert old_card_layout in text, "old_card_layout not found!"
text = text.replace(old_card_layout, new_card_layout, 1)

# 4. Update _list_card
old_list_card_prefix = """    @classmethod
    def _list_card(cls, card: CardData, x: float, y: float, slot_w: float, max_h: float,
                   d: Dict[str, Proto], size: float, bg_h: Optional[float] = None, idx: int = 0,
                   ncols: int = 1):
        \"\"\"Header + rows card (Library slides 1, 2 and 4), rows flowing into `ncols` columns.
        Without `bg_h` returns (required height, template-ratio height, max height) for the
        background container.\"\"\"
        decor, header, body, badge, divider = d["decor"], d["header"], d["body"], d.get("badge"), d["divider"]
        pad = T.CARD_PADDING

        # Section 5 geometric offsets and protrusion
        header_off_l = header.left - decor.left
        header_off_r = decor.right - header.right
        bg_off_top = decor.top - header.top
        body_off_l = body.left - decor.left
        body_off_r = decor.right - body.right
        badge_off_l = badge.left - decor.left if badge else body_off_l
        line_off_l = min(body_off_l, badge_off_l)
        ratio = decor.height / decor.width

        bg_w = slot_w
        if bg_off_top > 0:   # protruding header (Design 1)
            bg_top, header_top = y + bg_off_top, y
        else:                # flush header (Design 2 / CRD-SUB-LST)
            bg_top, header_top = y, y - bg_off_top
        max_bg_h = max_h - (bg_top - y)
        base_h = min(max(50.0, bg_w * ratio), max_bg_h)

        header_box = None
        start_y = bg_top + pad
        if card.header.strip():
            hw = bg_w - header_off_l - header_off_r
            hh = text_height(header, [card.header], hw, T.PILL_HEADER, bold=True, min_h=header.height)
            header_box = TextBox(_rect(x + header_off_l, header_top, hw, hh), [card.header], T.PILL_HEADER)
            start_y = header_top + hh + pad

        # Row columns share the inner width; the gap between them equals two body insets
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

        sep_h = divider.height
        columns = cls._split_columns(rows, ncols)

        def stack(col):
            return sum(r[5] for r in col) + max(0, len(col) - 1) * (2 * T.MIN_GAP + sep_h)

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
            empty = (end_y - start_y) - sum(r[5] for r in col) - (n - 1) * sep_h
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
                    row.divider = (col_x, cur, col_w)
                    cur += sep_h + gap_h
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
            header_style="floating_pill" if bg_off_top > 0 else "flush",
            rows=out_rows,
        )"""

new_list_card_prefix = """    @classmethod
    def _list_card(cls, card: CardData, x: float, y: float, slot_w: float, max_h: float,
                   d: Dict[str, Proto], size: float, bg_h: Optional[float] = None, idx: int = 0,
                   ncols: int = 1):
        \"\"\"Header + rows card (Library slides 1, 2, 4, 7 and 8), rows flowing into `ncols` columns.
        Without `bg_h` returns (required height, template-ratio height, max height) for the
        background container.\"\"\"
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

        sep_h = divider.height if has_row_divs else 0.0
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
        )"""

assert old_list_card_prefix in text, "old_list_card_prefix not found!"
text = text.replace(old_list_card_prefix, new_list_card_prefix, 1)

path.write_text(text, encoding="utf-8")
print("Successfully updated layout_engine.py!")

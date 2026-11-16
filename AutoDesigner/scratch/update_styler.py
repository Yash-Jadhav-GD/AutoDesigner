from pathlib import Path

path = Path("AutoDesigner/engine/styler.py")
text = path.read_text(encoding="utf-8")

# 1. Update imports
old_import = """try:
    from .analyzer import SlideData
    from .layout_engine import SlideLayout, CardLayout, Proto, Rect, TextBox
    from .builders.common import add_library_line, clean_element
    from . import svg_icons
except ImportError:
    from engine.analyzer import SlideData
    from engine.layout_engine import SlideLayout, CardLayout, Proto, Rect, TextBox
    from engine.builders.common import add_library_line, clean_element
    from engine import svg_icons"""

new_import = """try:
    from .analyzer import SlideData
    from .layout_engine import SlideLayout, CardLayout, Proto, Rect, TextBox
    from .builders.common import add_library_line, clean_element, set_shape_corner_radius
    from . import svg_icons
except ImportError:
    from engine.analyzer import SlideData
    from engine.layout_engine import SlideLayout, CardLayout, Proto, Rect, TextBox
    from engine.builders.common import add_library_line, clean_element, set_shape_corner_radius
    from engine import svg_icons"""

assert old_import in text, "old_import not found!"
text = text.replace(old_import, new_import, 1)

# 2. Update _render_card and add helpers
old_render_card = """    @classmethod
    def _render_card(cls, slide, card: CardLayout, design: Dict[str, Proto], palette: Dict[str, Any],
                     icons: Optional[List[svg_icons.IconRef]]):
        # 1. Background container (cloned first, so it sits behind the card's content)
        clone_proto(slide, design["decor"], card.container_rect)

        # 2. KPI number and header
        if card.number:
            cls._text(slide, design["number"], card.number)
        if card.header:
            header = cls._text(slide, design["header"], card.header)
            if card.header_style != "kpi":
                recolor_fill(header, palette.get("primary"))

        # 3. Rows: icon backplate + SVG icon, subheader, body, divider
        for row in card.rows:
            if row.badge_rect:
                clone_proto(slide, design["badge"], row.badge_rect)
                cls._icon(slide, row, icons)
            if row.sub:
                cls._text(slide, design["sub"], row.sub, autofit="spAutoFit")
            if row.text:
                cls._text(slide, design["body"], row.text, autofit="spAutoFit")
            if row.divider:
                left, top, width = row.divider
                line = add_library_line(slide, design["divider"].shape, left, top, width)
                _set_shape_id(line, slide)"""

new_render_card = """    @classmethod
    def _render_card(cls, slide, card: CardLayout, design: Dict[str, Proto], palette: Dict[str, Any],
                     icons: Optional[List[svg_icons.IconRef]]):
        # 1. Background container (cloned first, so it sits behind the card's content)
        container = clone_proto(slide, design["decor"], card.container_rect)
        set_shape_corner_radius(container, 4.0)

        # 2. KPI number and header
        if card.number:
            cls._text(slide, design["number"], card.number)

        # Header icon if present (crd_icn_v1)
        if card.header_icon_rect and "icon" in design:
            clone_proto(slide, design["icon"], card.header_icon_rect)
            if card.header_icon_text:
                cls._icon_in_rect(slide, card.header_icon_rect, card.header_icon_text, icons)

        if card.header:
            header = cls._text(slide, design["header"], card.header)
            if card.header_style not in ("kpi", "header_icon"):
                recolor_fill(header, palette.get("primary"))
            if card.header_style == "top_tab":
                set_shape_corner_radius(header, 4.0)

        # Header divider if present (crd_icn_v1)
        if card.header_divider and "divider" in design:
            left, top, width = card.header_divider
            line = add_library_line(slide, design["divider"].shape, left, top, width)
            _set_shape_id(line, slide)

        # 3. Rows: icon backplate + SVG icon, subheader, body, divider
        for row in card.rows:
            if row.badge_rect:
                clone_proto(slide, design["badge"], row.badge_rect)
                cls._icon(slide, row, icons)
            if row.sub:
                cls._text(slide, design["sub"], row.sub, autofit="spAutoFit")
            if row.text:
                txt_shape = cls._text(slide, design["body"], row.text, autofit="spAutoFit")
                if card.header_style == "top_tab":
                    cls._ensure_text_white(txt_shape)
            if row.divider:
                left, top, width = row.divider
                line = add_library_line(slide, design["divider"].shape, left, top, width)
                _set_shape_id(line, slide)

    @staticmethod
    def _icon_in_rect(slide, icon_rect: Rect, text: str, icons: Optional[List[svg_icons.IconRef]]):
        if icons is None or not icon_rect:
            return
        match = svg_icons.icon_for(text)
        if not match:
            return
        icon, svg = match
        rel_id = f"rIdGenIcon{len(icons) + 1}"
        icon_size = icon_rect.width * 0.6
        x = icon_rect.left + (icon_rect.width - icon_size) / 2
        y = icon_rect.top + (icon_rect.height - icon_size) / 2
        name = escape(f"Icon {icon.get('label', '')}", {'"': "&quot;"})
        slide.shapes._spTree.append(parse_xml(svg_icons.pic_xml(
            _next_shape_id(slide), name, rel_id, x, y, icon_size, icon_size)))
        icons.append((rel_id, svg_icons.media_name(icon), svg))

    @staticmethod
    def _ensure_text_white(shape):
        tx = shape._element.find(qn("p:txBody"))
        if tx is None:
            return
        for r_pr in tx.xpath(".//a:rPr"):
            for srgb in r_pr.xpath(".//a:srgbClr"):
                srgb.set("val", "FFFFFF")
            for scheme in r_pr.xpath(".//a:schemeClr"):
                scheme.set("val", "bg1")"""

assert old_render_card in text, "old_render_card not found!"
text = text.replace(old_render_card, new_render_card, 1)

path.write_text(text, encoding="utf-8")
print("Successfully updated styler.py!")

"""
AutoDesigner Styler
Renders the generic card archetypes by cloning the Library_copy.pptx prototype shapes
(fill, outline, corner radius, shadow, font, colour, alignment) at the positions computed by
the layout engine, with sizes from TypographyTokens and dividers from add_library_line.
"""

import copy
from xml.sax.saxutils import escape
from typing import Dict, Any, List, Optional

from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn

try:
    from .analyzer import SlideData
    from .layout_engine import SlideLayout, CardLayout, Proto, Rect, TextBox
    from .builders.common import add_library_line, clean_element, set_shape_corner_radius
    from . import svg_icons
except ImportError:
    from engine.analyzer import SlideData
    from engine.layout_engine import SlideLayout, CardLayout, Proto, Rect, TextBox
    from engine.builders.common import add_library_line, clean_element, set_shape_corner_radius
    from engine import svg_icons


def _next_shape_id(slide) -> int:
    ids = [int(e.get("id")) for e in slide.shapes._spTree.iter(qn("p:cNvPr")) if e.get("id", "").isdigit()]
    return max(ids, default=1) + 1


def _set_shape_id(shape, slide):
    cnv = shape._element.find(".//" + qn("p:cNvPr"))
    for ext in cnv.findall(qn("a:extLst")):
        cnv.remove(ext)
    cnv.set("id", str(_next_shape_id(slide)))


def clone_proto(slide, proto: Proto, rect: Rect):
    """Appends a copy of a Library prototype shape at `rect`, with a fresh shape id."""
    elem = clean_element(copy.deepcopy(proto.shape._element))
    slide.shapes._spTree.append(elem)
    shape = slide.shapes[len(slide.shapes) - 1]
    _set_shape_id(shape, slide)
    shape.left, shape.top, shape.width, shape.height = rect.left, rect.top, rect.width, rect.height
    return shape


def set_text(shape, paragraphs: List[str], size_pt: float, autofit: Optional[str] = None):
    """Replaces the prototype's text, keeping its paragraph and run formatting; sets the size."""
    tx = shape._element.find(qn("p:txBody"))
    body_pr = tx.find(qn("a:bodyPr"))
    body_pr.set("wrap", "square")
    if autofit:
        for tag in ("a:noAutofit", "a:normAutofit", "a:spAutoFit"):
            for old in body_pr.findall(qn(tag)):
                body_pr.remove(old)
        warp = body_pr.find(qn("a:prstTxWarp"))
        fit = parse_xml(f'<a:{autofit} xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>')
        if warp is not None:
            warp.addnext(fit)
        else:
            body_pr.insert(0, fit)

    old_ps = tx.findall(qn("a:p"))
    p0 = old_ps[0]
    p_pr = p0.find(qn("a:pPr"))
    run = p0.find(qn("a:r"))
    r_pr = run.find(qn("a:rPr")) if run is not None else p0.find(qn("a:endParaRPr"))
    for p in old_ps:
        tx.remove(p)

    sz = str(int(round(size_pt * 100)))
    for text in paragraphs:
        p = tx.makeelement(qn("a:p"), {})
        if p_pr is not None:
            p.append(copy.deepcopy(p_pr))
        r = p.makeelement(qn("a:r"), {})
        rp = copy.deepcopy(r_pr) if r_pr is not None else r.makeelement(qn("a:rPr"), {})
        rp.tag = qn("a:rPr")
        rp.set("sz", sz)
        rp.set("dirty", "0")
        r.append(rp)
        t = r.makeelement(qn("a:t"), {})
        t.text = text
        r.append(t)
        p.append(r)
        tx.append(p)


def recolor_fill(shape, hex_val: str):
    """Swaps the prototype's solid fill colour (used to apply the palette's primary to headers)."""
    clr = shape._element.find(f"{qn('p:spPr')}/{qn('a:solidFill')}/{qn('a:srgbClr')}")
    if clr is not None and hex_val:
        clr.set("val", hex_val.lstrip("#").upper())


class SlideStyler:
    """Renders styled cards, titles, headers, and dividers on a presentation slide."""

    FONT_FAMILY = "Aptos"
    FALLBACK_FONT = "Segoe UI"

    @classmethod
    def apply_style(
        cls,
        slide,
        slide_data: SlideData,
        layout: SlideLayout,
        palette: Dict[str, Any],
        icons: Optional[List[svg_icons.IconRef]] = None,
    ):
        """`icons` collects the SVGs placed on icon_badge rows, for svg_icons.embed_svg_icons."""
        # 1. Background fill
        cls._apply_slide_background(slide, palette)

        # 2. Slide Title & Category
        cls._render_title_block(slide, slide_data, layout, palette)

        # 3. Render Cards
        for card_layout in layout.cards:
            cls._render_card(slide, card_layout, layout.design, palette, icons)

    @classmethod
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
                scheme.set("val", "bg1")

    @staticmethod
    def _text(slide, proto: Proto, box: TextBox, autofit: Optional[str] = None):
        shape = clone_proto(slide, proto, box.rect)
        set_text(shape, box.paragraphs, box.size, autofit)
        return shape

    @staticmethod
    def _icon(slide, row, icons: Optional[List[svg_icons.IconRef]]):
        """Best keyword-matched SVG centred in the backplate; the backplate stays empty without one."""
        if icons is None or not row.icon_rect:
            return
        match = svg_icons.icon_for(row.icon_text)
        if not match:
            return
        icon, svg = match
        rel_id = f"rIdGenIcon{len(icons) + 1}"
        r = row.icon_rect
        name = escape(f"Icon {icon.get('label', '')}", {'"': "&quot;"})
        slide.shapes._spTree.append(parse_xml(svg_icons.pic_xml(
            _next_shape_id(slide), name, rel_id, r.left, r.top, r.width, r.height)))
        icons.append((rel_id, svg_icons.media_name(icon), svg))

    @classmethod
    def _apply_slide_background(cls, slide, palette: Dict[str, Any]):
        bg = slide.background
        fill = bg.fill
        fill.solid()
        bg_hex = palette.get("slide_bg", "F8FAFC")
        fill.fore_color.rgb = RGBColor.from_string(bg_hex)

    @classmethod
    def _render_title_block(
        cls, slide, slide_data: SlideData, layout: SlideLayout, palette: Dict[str, Any]
    ):
        if not layout.title_rect or not slide_data.title or not slide_data.title.strip():
            return
        # Category pill if available
        if slide_data.category and layout.category_rect:
            cat_shape = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                layout.category_rect.left,
                layout.category_rect.top,
                layout.category_rect.width,
                layout.category_rect.height,
            )
            cat_shape.fill.solid()
            cat_shape.fill.fore_color.rgb = RGBColor.from_string(palette.get("badge_bg", "F1EEF8"))
            cat_shape.line.fill.background()
            tf = cat_shape.text_frame
            tf.word_wrap = False
            tf.margin_top = Pt(1)
            tf.margin_bottom = Pt(1)
            tf.margin_left = Pt(8)
            tf.margin_right = Pt(8)
            p = tf.paragraphs[0]
            p.text = slide_data.category.upper()
            p.alignment = PP_ALIGN.CENTER
            p.font.name = cls.FONT_FAMILY
            p.font.size = Pt(9.5)
            p.font.bold = True
            p.font.color.rgb = RGBColor.from_string(palette.get("primary", "48397E"))

        # Main Title
        title_box = slide.shapes.add_textbox(
            layout.title_rect.left,
            layout.title_rect.top,
            layout.title_rect.width,
            layout.title_rect.height,
        )
        tf = title_box.text_frame
        tf.word_wrap = True
        tf.margin_left = Pt(0)
        tf.margin_right = Pt(0)
        tf.margin_top = Pt(0)
        tf.margin_bottom = Pt(0)
        p = tf.paragraphs[0]
        p.text = slide_data.title if slide_data.title else "Executive Summary"
        p.font.name = cls.FONT_FAMILY
        p.font.size = Pt(22)
        p.font.bold = True
        p.font.color.rgb = RGBColor.from_string(palette.get("text_dark", "262626"))

        # Subtitle
        if slide_data.subtitle and layout.subtitle_rect:
            p2 = tf.add_paragraph()
            p2.text = slide_data.subtitle
            p2.font.name = cls.FONT_FAMILY
            p2.font.size = Pt(12)
            p2.font.color.rgb = RGBColor.from_string(palette.get("text_muted", "64748B"))

        # Accent divider line beneath title
        divider_y = layout.title_rect.top + layout.title_rect.height + Pt(4)
        div = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            layout.title_rect.left,
            divider_y,
            layout.title_rect.width,
            Pt(1),
        )
        div.fill.solid()
        div.fill.fore_color.rgb = RGBColor.from_string(palette.get("divider", "E2E8F0"))
        div.line.fill.background()


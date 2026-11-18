"""
Universal DrawingML Composable Renderer
Renders ResolvedShape objects onto a PowerPoint slide.
Enforces 100% middle vertical alignment, exact 4.0pt corner radius,
DrawingML native bullets, authentic Library divider lines, and clean fills.
"""

import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import pptx
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml import parse_xml

from ..color_tokens import (
    ColorTokens,
    apply_solid_fill,
    apply_linear_gradient,
    hex_to_rgb,
)
from ..builders.common import (
    set_shape_corner_radius,
    apply_native_bullet,
    clean_bullet_string,
    add_library_line,
    apply_library_shadow,
    set_cell_border,
)
from .layout_solver import ResolvedShape


def resolve_color(token_or_hex: Optional[str]) -> str:
    """Resolves a color token name or hex string to a clean 6-char hex code."""
    if not token_or_hex:
        return ColorTokens.NEUTRAL_WHITE
    token_map = {
        "BRAND_INDIGO": ColorTokens.BRAND_INDIGO,
        "BRAND_TANGERINE": ColorTokens.BRAND_TANGERINE,
        "BRAND_NAVY": ColorTokens.BRAND_NAVY,
        "BRAND_ONYX": ColorTokens.BRAND_ONYX,
        "VIVID_INDIGO_MID": ColorTokens.VIVID_INDIGO_MID,
        "VIVID_INDIGO_SOFT": ColorTokens.VIVID_INDIGO_SOFT,
        "VIVID_INDIGO_WASH": ColorTokens.VIVID_INDIGO_WASH,
        "VIVID_INDIGO_LIGHT": ColorTokens.VIVID_INDIGO_LIGHT,
        "VIVID_NAVY_MID": ColorTokens.VIVID_NAVY_MID,
        "VIVID_NAVY_SOFT": ColorTokens.VIVID_NAVY_SOFT,
        "VIVID_TANGERINE_MID": ColorTokens.VIVID_TANGERINE_MID,
        "VIVID_TANGERINE_SOFT": ColorTokens.VIVID_TANGERINE_SOFT,
        "NEUTRAL_WHITE": ColorTokens.NEUTRAL_WHITE,
        "NEUTRAL_SURFACE": ColorTokens.NEUTRAL_SURFACE,
        "NEUTRAL_BORDER": ColorTokens.NEUTRAL_BORDER,
        "NEUTRAL_LINE": ColorTokens.NEUTRAL_LINE,
    }
    cleaned = token_or_hex.strip().lstrip("#")
    return token_map.get(token_or_hex, cleaned)


def apply_solid_fill_alpha(shape, hex_val: str, alpha_val: int = 100000):
    """Sets DrawingML solidFill with optional alpha (transparency: 100000 = 100%, 20000 = 20%)."""
    spPr = shape._element.spPr
    for tag in ['solidFill', 'gradFill', 'blipFill', 'pattFill', 'noFill']:
        elem = spPr.find(f'{{http://schemas.openxmlformats.org/drawingml/2006/main}}{tag}')
        if elem is not None:
            spPr.remove(elem)
    hex_clean = hex_val.lstrip('#')
    if alpha_val < 100000:
        fill_xml = f'''<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:srgbClr val="{hex_clean}">
            <a:alpha val="{alpha_val}"/>
          </a:srgbClr>
        </a:solidFill>'''
    else:
        fill_xml = f'''<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:srgbClr val="{hex_clean}"/>
        </a:solidFill>'''
    fill_elem = parse_xml(fill_xml)
    ln = spPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}ln')
    if ln is not None:
        ln_idx = list(spPr).index(ln)
        spPr.insert(ln_idx, fill_elem)
    else:
        spPr.append(fill_elem)


def pic_xml(shape_id: int, name: str, rel_id: str, left_pt: float, top_pt: float, width_pt: float, height_pt: float) -> str:
    """svgBlip picture element in DrawingML."""
    l_emu = int(left_pt * 12700)
    t_emu = int(top_pt * 12700)
    w_emu = int(width_pt * 12700)
    h_emu = int(height_pt * 12700)
    return f'''<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:nvPicPr>
        <p:cNvPr id="{shape_id}" name="{name}"/>
        <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
        <p:nvPr/>
      </p:nvPicPr>
      <p:blipFill>
        <a:blip>
          <a:extLst>
            <a:ext uri="{{96DAC541-7B7A-43D3-8B79-37D633B846F1}}">
              <asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="{rel_id}"/>
            </a:ext>
          </a:extLst>
        </a:blip>
        <a:stretch><a:fillRect/></a:stretch>
      </p:blipFill>
      <p:spPr>
        <a:xfrm>
          <a:off x="{l_emu}" y="{t_emu}"/>
          <a:ext cx="{w_emu}" cy="{h_emu}"/>
        </a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </p:spPr>
    </p:pic>'''


def load_svg_bytes(asset_name: Optional[str], fallback_bytes: Optional[bytes] = None) -> Optional[bytes]:
    """Retrieves SVG icon bytes from library assets, local icon cache, or explicit bytes."""
    if fallback_bytes:
        return fallback_bytes
    if not asset_name:
        return None
    p_lib = Path(__file__).resolve().parent.parent / "assets" / "library_icons" / asset_name
    if p_lib.exists():
        return p_lib.read_bytes()
    p_cache = Path(__file__).resolve().parents[2] / ".cache" / "icons" / asset_name
    if p_cache.exists():
        return p_cache.read_bytes()
    return None


class ComposableRenderer:
    """Renders resolved shapes onto PowerPoint slides."""

    @classmethod
    def render(
        cls,
        slide,
        resolved_shapes: List[ResolvedShape],
        lib_lines: Dict[str, Any],
        lib_shadow_proto: Optional[Any] = None
    ) -> List[Tuple[str, str, bytes]]:
        """Renders all resolved shapes onto slide sorted by z-order."""
        slide_icon_refs: List[Tuple[str, str, bytes]] = []
        # Sort by z-order so containers sit underneath text and headers
        sorted_shapes = sorted(resolved_shapes, key=lambda s: s.z_order)

        for sh in sorted_shapes:
            r = sh.rect
            kind = sh.kind

            if kind in ("card_container", "kpi_container", "process_row"):
                box = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    Pt(r.left), Pt(r.top), Pt(r.width), Pt(r.height)
                )
                box.name = sh.shape_id
                set_shape_corner_radius(box, sh.corner_radius)
                fill_hex = resolve_color(sh.fill_color)
                alpha_val = sh.metadata.get("alpha")
                if alpha_val is not None:
                    apply_solid_fill_alpha(box, fill_hex, int(alpha_val))
                else:
                    apply_solid_fill(box, fill_hex)

                if sh.border_color:
                    box.line.color.rgb = hex_to_rgb(resolve_color(sh.border_color))
                    box.line.width = Pt(sh.border_width)
                else:
                    box.line.fill.background()

                # Rule: DO NOT ADD SHADOWS ANYWHERE

            elif kind in ("header_pill", "flush_header"):
                pill = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    Pt(r.left), Pt(r.top), Pt(r.width), Pt(r.height)
                )
                pill.name = sh.shape_id
                set_shape_corner_radius(pill, sh.corner_radius)

                # Gradient or Solid Fill (Zero shadows everywhere)
                if sh.gradient == "linear_indigo":
                    apply_linear_gradient(pill, ColorTokens.GRADIENT_INDIGO_START, ColorTokens.GRADIENT_INDIGO_END, with_shadow=False)
                elif sh.gradient == "linear_navy":
                    apply_linear_gradient(pill, ColorTokens.GRADIENT_NAVY_START, ColorTokens.GRADIENT_NAVY_END, with_shadow=False)
                elif sh.gradient == "linear_tangerine":
                    apply_linear_gradient(pill, ColorTokens.GRADIENT_TANGERINE_START, ColorTokens.GRADIENT_TANGERINE_END, with_shadow=False)
                else:
                    apply_solid_fill(pill, resolve_color(sh.fill_color))

                pill.line.fill.background()

                # Text in Pill
                tf = pill.text_frame
                tf.word_wrap = True
                tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                tf._txBody.bodyPr.set("anchor", "ctr")
                tf.text = sh.text or ""
                p = tf.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER if sh.align == "center" else PP_ALIGN.LEFT
                p.font.name = "Aptos"
                p.font.size = Pt(sh.font_size)
                p.font.bold = sh.bold
                p.font.color.rgb = hex_to_rgb(resolve_color(sh.text_color))

            elif sh.paragraphs or kind == "kpi_content":
                tb = slide.shapes.add_textbox(Pt(r.left), Pt(r.top), Pt(r.width), Pt(r.height))
                tb.name = sh.shape_id
                tf = tb.text_frame
                tf.word_wrap = True
                tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                tf._txBody.bodyPr.set("anchor", "ctr")
                paras = sh.paragraphs or []
                first = True
                for p_data in paras:
                    p = tf.paragraphs[0] if first else tf.add_paragraph()
                    first = False
                    p.text = p_data.get("text", "")
                    p.alignment = PP_ALIGN.CENTER if p_data.get("align", sh.align) == "center" else PP_ALIGN.LEFT
                    p.font.name = "Aptos"
                    p.font.size = Pt(p_data.get("font_size", sh.font_size))
                    p.font.bold = p_data.get("bold", False)
                    p_col = resolve_color(p_data.get("color", sh.text_color))
                    p.font.color.rgb = hex_to_rgb(p_col)
                    if "space_after" in p_data:
                        p.space_after = Pt(p_data["space_after"])
                    if "space_before" in p_data:
                        p.space_before = Pt(p_data["space_before"])

            elif kind in ("text_box", "kpi_text"):
                tb = slide.shapes.add_textbox(Pt(r.left), Pt(r.top), Pt(r.width), Pt(r.height))
                tb.name = sh.shape_id
                tf = tb.text_frame
                tf.word_wrap = True
                tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                tf._txBody.bodyPr.set("anchor", "ctr")
                tf.text = sh.text or ""
                for p in tf.paragraphs:
                    p.alignment = PP_ALIGN.CENTER if sh.align == "center" else PP_ALIGN.LEFT
                    p.font.name = "Aptos"
                    p.font.size = Pt(sh.font_size)
                    p.font.bold = sh.bold
                    p.font.color.rgb = hex_to_rgb(resolve_color(sh.text_color))

            elif kind == "bullet_list":
                tb = slide.shapes.add_textbox(Pt(r.left), Pt(r.top), Pt(r.width), Pt(r.height))
                tb.name = sh.shape_id
                tf = tb.text_frame
                tf.word_wrap = True
                tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                tf._txBody.bodyPr.set("anchor", "ctr")

                cfg = sh.bullet_config or {}
                items = cfg.get("items", [])
                space_after = cfg.get("space_after", 7.0)
                bullet_char = cfg.get("bullet_char", "•")
                mar_l = cfg.get("mar_l_pt", 13.5)
                indent = cfg.get("indent_pt", -13.5)

                for idx, item in enumerate(items):
                    p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
                    clean_txt = clean_bullet_string(item)
                    p.text = clean_txt
                    p.font.name = "Aptos"
                    p.font.size = Pt(sh.font_size)
                    p.font.bold = False
                    p.font.color.rgb = hex_to_rgb(resolve_color(sh.text_color))
                    p.space_after = Pt(space_after)
                    if bullet_char:
                        apply_native_bullet(p, char=bullet_char, typeface="Arial", mar_l_pt=mar_l, indent_pt=indent)
                    else:
                        p.alignment = PP_ALIGN.CENTER if sh.align == "center" else PP_ALIGN.LEFT

            elif kind == "section_list":
                tb = slide.shapes.add_textbox(Pt(r.left), Pt(r.top), Pt(r.width), Pt(r.height))
                tb.name = sh.shape_id
                tf = tb.text_frame
                tf.word_wrap = True
                tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                tf._txBody.bodyPr.set("anchor", "ctr")

                sections = sh.metadata.get("sections", [])
                h_size = sh.metadata.get("heading_font_size", 11.5)
                h_color = resolve_color(sh.metadata.get("heading_color", "BRAND_INDIGO"))
                b_size = sh.metadata.get("bullet_font_size", 10.5)
                b_color = resolve_color(sh.metadata.get("bullet_color", "BRAND_ONYX"))
                sec_gap = sh.metadata.get("space_after_section", 10.0)

                first_para = True
                for sec in sections:
                    heading = sec.get("heading", "")
                    p_head = tf.paragraphs[0] if first_para else tf.add_paragraph()
                    first_para = False
                    p_head.text = heading
                    p_head.font.name = "Aptos"
                    p_head.font.size = Pt(h_size)
                    p_head.font.bold = True
                    p_head.font.color.rgb = hex_to_rgb(h_color)
                    p_head.space_after = Pt(3.0)

                    bullets = sec.get("bullets", [])
                    for b_idx, bullet in enumerate(bullets):
                        p_b = tf.add_paragraph()
                        p_b.text = clean_bullet_string(bullet)
                        p_b.font.name = "Aptos"
                        p_b.font.size = Pt(b_size)
                        p_b.font.bold = False
                        p_b.font.color.rgb = hex_to_rgb(b_color)
                        p_b.space_after = Pt(sec_gap if b_idx == len(bullets) - 1 else 4.0)
                        apply_native_bullet(p_b, char="•", typeface="Arial", mar_l_pt=11.0, indent_pt=-11.0)

            elif kind == "divider":
                proto_key = sh.line_prototype or "side_fade_r"
                proto = lib_lines.get(proto_key)
                if proto is not None:
                    line = add_library_line(slide, proto, r.left, r.top, r.width, r.height)
                    line.name = sh.shape_id

            elif kind == "table":
                tdata = sh.table_data or {}
                headers = tdata.get("headers", [])
                rows = tdata.get("rows", [])
                col_widths = tdata.get("col_widths")
                h_fill = resolve_color(tdata.get("header_fill", "BRAND_INDIGO"))
                z_fill = resolve_color(tdata.get("zebra_fill", "NEUTRAL_SURFACE"))
                tot_fill = resolve_color(tdata.get("totals_fill", "BRAND_NAVY"))
                has_totals = tdata.get("totals_row", False)
                font_sz = tdata.get("font_size", 9.5)

                num_rows = len(rows) + (1 if headers else 0)
                num_cols = len(headers) if headers else (len(rows[0]) if rows else 1)

                table_shape = slide.shapes.add_table(num_rows, num_cols, Pt(r.left), Pt(r.top), Pt(r.width), Pt(r.height))
                table_shape.name = sh.shape_id
                tbl = table_shape.table

                # Set custom col widths if provided
                if col_widths and len(col_widths) == num_cols:
                    for c_idx, w in enumerate(col_widths):
                        tbl.columns[c_idx].width = Pt(w)

                row_offset = 0
                if headers:
                    for c_idx, h_text in enumerate(headers):
                        cell = tbl.cell(0, c_idx)
                        cell.text = h_text
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = hex_to_rgb(h_fill)
                        set_cell_border(cell, color="FFFFFF", width="6350")
                        for p in cell.text_frame.paragraphs:
                            p.alignment = PP_ALIGN.CENTER
                            p.font.name = "Aptos"
                            p.font.size = Pt(font_sz)
                            p.font.bold = True
                            p.font.color.rgb = hex_to_rgb(ColorTokens.TEXT_ON_DARK)
                    row_offset = 1

                for r_idx, row_items in enumerate(rows):
                    is_last_totals = has_totals and (r_idx == len(rows) - 1)
                    curr_fill = tot_fill if is_last_totals else (z_fill if r_idx % 2 == 1 else ColorTokens.NEUTRAL_WHITE)
                    for c_idx, cell_value in enumerate(row_items):
                        if c_idx < num_cols:
                            cell = tbl.cell(r_idx + row_offset, c_idx)
                            cell.text = str(cell_value)
                            cell.fill.solid()
                            cell.fill.fore_color.rgb = hex_to_rgb(curr_fill)
                            set_cell_border(cell, color="CBD5E1", width="6350")
                            for p in cell.text_frame.paragraphs:
                                p.alignment = PP_ALIGN.LEFT if c_idx == 0 else PP_ALIGN.RIGHT
                                p.font.name = "Aptos"
                                p.font.size = Pt(font_sz)
                                p.font.bold = is_last_totals
                                p.font.color.rgb = hex_to_rgb(ColorTokens.TEXT_ON_DARK if is_last_totals else ColorTokens.TEXT_PRIMARY)

            elif kind == "badge_oval":
                oval = slide.shapes.add_shape(
                    MSO_SHAPE.OVAL,
                    Pt(r.left), Pt(r.top), Pt(r.width), Pt(r.height)
                )
                oval.name = sh.shape_id
                fill_hex = resolve_color(sh.fill_color)
                apply_solid_fill(oval, fill_hex)
                if sh.border_color:
                    oval.line.color.rgb = hex_to_rgb(resolve_color(sh.border_color))
                    oval.line.width = Pt(sh.border_width)
                else:
                    oval.line.fill.background()

                tf = oval.text_frame
                tf.word_wrap = False
                tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                tf._txBody.bodyPr.set("anchor", "ctr")
                tf.text = sh.text or ""
                p = tf.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER
                p.font.name = "Aptos"
                p.font.size = Pt(sh.font_size)
                p.font.bold = sh.bold
                if sh.metadata.get("italic"):
                    p.font.italic = True
                p.font.color.rgb = hex_to_rgb(resolve_color(sh.text_color))

            elif kind in ("solid_line", "vertical_line"):
                line_color = resolve_color(sh.border_color or sh.fill_color)
                line = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE,
                    Pt(r.left), Pt(r.top), Pt(r.width), Pt(r.height)
                )
                line.name = sh.shape_id
                apply_solid_fill(line, line_color)
                line.line.color.rgb = hex_to_rgb(line_color)
                line.line.width = Pt(sh.border_width)

            elif kind == "svg_icon":
                svg_data = load_svg_bytes(sh.svg_asset, sh.metadata.get("svg_bytes"))
                if svg_data:
                    # Apply color tint if specified
                    tint_hex = sh.metadata.get("tint")
                    if tint_hex:
                        clean_tint = "#" + resolve_color(tint_hex).lstrip("#")
                        text = svg_data.decode("utf-8", errors="ignore")
                        text = re.sub(r"#66737[cC]", clean_tint, text)
                        text = text.replace("currentColor", clean_tint)
                        svg_data = text.encode("utf-8")
                    rel_id = f"rIdSvg_{len(slide_icon_refs) + 100}"
                    media_name = f"icon_{sh.shape_id}_{sh.svg_asset or 'asset.svg'}"
                    slide_icon_refs.append((rel_id, media_name, svg_data))
                    shape_numeric_id = 500 + len(slide_icon_refs)
                    pic_element = parse_xml(pic_xml(
                        shape_numeric_id,
                        sh.shape_id,
                        rel_id,
                        r.left, r.top, r.width, r.height
                    ))
                    slide.shapes._spTree.append(pic_element)

        return slide_icon_refs

"""
Universal DrawingML Composable Renderer
Renders ResolvedShape objects onto a PowerPoint slide.
Enforces 100% middle vertical alignment, exact 4.0pt corner radius,
DrawingML native bullets, authentic Library divider lines, and clean fills.
"""

from typing import Any, Dict, List, Optional
import pptx
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN

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


class ComposableRenderer:
    """Renders resolved shapes onto PowerPoint slides."""

    @classmethod
    def render(
        cls,
        slide,
        resolved_shapes: List[ResolvedShape],
        lib_lines: Dict[str, Any],
        lib_shadow_proto: Optional[Any] = None
    ):
        """Renders all resolved shapes onto slide sorted by z-order."""
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
                apply_solid_fill(box, fill_hex)

                if sh.border_color:
                    box.line.color.rgb = hex_to_rgb(resolve_color(sh.border_color))
                    box.line.width = Pt(sh.border_width)
                else:
                    box.line.fill.background()

                if sh.shadow and lib_shadow_proto:
                    apply_library_shadow(box, lib_shadow_proto)

            elif kind in ("header_pill", "flush_header"):
                pill = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    Pt(r.left), Pt(r.top), Pt(r.width), Pt(r.height)
                )
                pill.name = sh.shape_id
                set_shape_corner_radius(pill, sh.corner_radius)

                # Gradient or Solid Fill
                if sh.gradient == "linear_indigo":
                    apply_linear_gradient(pill, ColorTokens.GRADIENT_INDIGO_START, ColorTokens.GRADIENT_INDIGO_END, with_shadow=sh.shadow)
                elif sh.gradient == "linear_navy":
                    apply_linear_gradient(pill, ColorTokens.GRADIENT_NAVY_START, ColorTokens.GRADIENT_NAVY_END, with_shadow=sh.shadow)
                elif sh.gradient == "linear_tangerine":
                    apply_linear_gradient(pill, ColorTokens.GRADIENT_TANGERINE_START, ColorTokens.GRADIENT_TANGERINE_END, with_shadow=sh.shadow)
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

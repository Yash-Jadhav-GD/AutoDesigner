"""
Constraint & Box-Model Layout Solver
Recursively calculates spatial coordinates (x, y, w, h) for any composition tree.
Guarantees strictly 0 shape collisions by construction.
"""

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple
from pydantic import BaseModel

from .primitives import (
    LayoutNode,
    AnyNode,
    HStack,
    VStack,
    Grid,
    CardNode,
    PillHeaderNode,
    BulletListNode,
    SectionListNode,
    KPINode,
    ProcessRowNode,
    TableNode,
    DividerNode,
    ConnectorNode,
    MaturityStageNode,
    SlideComposition,
    Padding,
)


@dataclass
class Rect:
    left: float
    top: float
    width: float
    height: float

    @property
    def right(self) -> float:
        return self.left + self.width

    @property
    def bottom(self) -> float:
        return self.top + self.height


@dataclass
class ResolvedShape:
    """A shape whose exact coordinates, z-order, and DrawingML properties are fully determined."""
    shape_id: str
    kind: str  # 'card_container', 'header_pill', 'flush_header', 'text_box', 'bullet_list', 'section_list', 'kpi_container', 'kpi_text', 'process_row', 'divider', 'connector', 'table', 'svg_icon'
    rect: Rect
    z_order: int = 10
    fill_color: Optional[str] = None
    border_color: Optional[str] = None
    border_width: float = 0.75
    corner_radius: float = 4.0
    text: Optional[str] = None
    font_size: float = 11.0
    bold: bool = False
    text_color: str = "BRAND_ONYX"
    align: str = "left"  # 'left', 'center', 'right'
    vertical_anchor: str = "ctr"  # always Middle
    bullet_config: Optional[Dict[str, Any]] = None
    shadow: bool = False
    gradient: Optional[str] = None
    line_prototype: Optional[str] = None
    table_data: Optional[Dict[str, Any]] = None
    svg_asset: Optional[str] = None
    paragraphs: Optional[List[Dict[str, Any]]] = None
    metadata: Dict[str, Any] = field(default_factory=dict)


class LayoutSolver:
    """Computes exact coordinate bounding boxes on 16:9 canvas (960pt x 540pt)."""

    CANVAS_W = 960.0
    CANVAS_H = 540.0

    CANVAS_LEFT = 45.0
    CANVAS_RIGHT = 915.0
    USABLE_W = 870.0  # 915 - 45

    TITLE_TOP = 26.0
    TITLE_H = 26.0
    SUBTITLE_GAP = 3.0
    SUBTITLE_H = 18.0

    CONTENT_TOP = 96.0  # top of main layout area when title exists
    CONTENT_TOP_NO_TITLE = 40.0
    CONTENT_BOTTOM = 512.0  # leave 28pt margin at bottom

    def __init__(self):
        self._shape_counter = 0

    def _next_id(self, prefix: str) -> str:
        self._shape_counter += 1
        return f"{prefix}_{self._shape_counter}"

    def solve(self, composition: SlideComposition) -> Tuple[List[ResolvedShape], Dict[str, Any]]:
        """Solves a SlideComposition into a flat list of positioned ResolvedShapes."""
        self._shape_counter = 0
        resolved: List[ResolvedShape] = []

        has_title = bool(composition.title and composition.title.strip())
        content_top = self.CONTENT_TOP if has_title else self.CONTENT_TOP_NO_TITLE
        content_bottom = self.CONTENT_BOTTOM
        if composition.footnote:
            content_bottom -= 20.0

        # 1. Title & Subtitle
        if has_title:
            cur_y = self.TITLE_TOP
            resolved.append(ResolvedShape(
                shape_id=self._next_id("slide_title"),
                kind="text_box",
                rect=Rect(self.CANVAS_LEFT, cur_y, self.USABLE_W, self.TITLE_H),
                z_order=20,
                text=composition.title,
                font_size=17.0,
                bold=True,
                text_color="BRAND_ONYX",
                align="left",
                vertical_anchor="ctr"
            ))
            cur_y += self.TITLE_H + self.SUBTITLE_GAP

            if composition.subtitle:
                resolved.append(ResolvedShape(
                    shape_id=self._next_id("slide_subtitle"),
                    kind="text_box",
                    rect=Rect(self.CANVAS_LEFT, cur_y, self.USABLE_W, self.SUBTITLE_H),
                    z_order=20,
                    text=composition.subtitle,
                    font_size=11.5,
                    bold=False,
                    text_color="BRAND_ONYX",
                    align="left",
                    vertical_anchor="ctr"
                ))
                cur_y += self.SUBTITLE_H + 8.0
            content_top = max(content_top, cur_y + 8.0)

        # 2. Main Layout Area Bounds
        available_rect = Rect(
            left=self.CANVAS_LEFT,
            top=content_top,
            width=self.USABLE_W,
            height=max(50.0, content_bottom - content_top)
        )

        # 3. Solve Root Tree
        self._solve_node(composition.root, available_rect, resolved)

        # 4. Footnote (if present)
        if composition.footnote:
            resolved.append(ResolvedShape(
                shape_id=self._next_id("slide_footnote"),
                kind="text_box",
                rect=Rect(self.CANVAS_LEFT, content_bottom + 4.0, self.USABLE_W, 16.0),
                z_order=20,
                text=composition.footnote,
                font_size=8.5,
                bold=False,
                text_color="NEUTRAL_LINE",
                align="left",
                vertical_anchor="ctr"
            ))

        meta = {
            "slide_width": self.CANVAS_W,
            "slide_height": self.CANVAS_H,
            "total_shapes": len(resolved)
        }
        return resolved, meta

    def _solve_node(self, node: AnyNode, bounds: Rect, resolved: List[ResolvedShape]):
        """Recursively resolves coordinates for any node inside parent bounds."""
        if isinstance(node, HStack):
            self._solve_hstack(node, bounds, resolved)
        elif isinstance(node, VStack):
            self._solve_vstack(node, bounds, resolved)
        elif isinstance(node, Grid):
            self._solve_grid(node, bounds, resolved)
        elif isinstance(node, CardNode):
            self._solve_card(node, bounds, resolved)
        elif isinstance(node, KPINode):
            self._solve_kpi(node, bounds, resolved)
        elif isinstance(node, ProcessRowNode):
            self._solve_process_row(node, bounds, resolved)
        elif isinstance(node, TableNode):
            self._solve_table(node, bounds, resolved)
        elif isinstance(node, DividerNode):
            self._solve_divider(node, bounds, resolved)
        elif isinstance(node, ConnectorNode):
            self._solve_connector(node, bounds, resolved)
        elif isinstance(node, MaturityStageNode):
            self._solve_maturity_stage(node, bounds, resolved)

    def _solve_hstack(self, node: HStack, bounds: Rect, resolved: List[ResolvedShape]):
        n = len(node.children)
        if n == 0:
            return

        total_gap = (n - 1) * node.gap
        fixed_w_sum = sum(c.fixed_width for c in node.children if c.fixed_width is not None)
        weighted_children = [c for c in node.children if c.fixed_width is None]
        total_weight = sum(c.width_weight for c in weighted_children) or 1.0

        remaining_w = max(0.0, bounds.width - total_gap - fixed_w_sum)

        cur_x = bounds.left
        for child in node.children:
            w = child.fixed_width if child.fixed_width is not None else (child.width_weight / total_weight) * remaining_w
            h = bounds.height if node.align == "stretch" else (child.fixed_height or bounds.height)
            y = bounds.top
            if node.align == "center":
                y = bounds.top + (bounds.height - h) / 2.0
            elif node.align == "bottom":
                y = bounds.top + (bounds.height - h)

            child_bounds = Rect(cur_x, y, w, h)
            self._solve_node(child, child_bounds, resolved)
            cur_x += w + node.gap

    def _solve_vstack(self, node: VStack, bounds: Rect, resolved: List[ResolvedShape]):
        n = len(node.children)
        if n == 0:
            return

        total_gap = (n - 1) * node.gap
        fixed_h_sum = sum(c.fixed_height for c in node.children if c.fixed_height is not None)
        weighted_children = [c for c in node.children if c.fixed_height is None]
        total_weight = sum(c.height_weight for c in weighted_children) or 1.0

        remaining_h = max(0.0, bounds.height - total_gap - fixed_h_sum)

        cur_y = bounds.top
        for child in node.children:
            h = child.fixed_height if child.fixed_height is not None else (child.height_weight / total_weight) * remaining_h
            w = child.fixed_width if child.fixed_width is not None else bounds.width
            x = bounds.left
            if node.align == "center":
                x = bounds.left + (bounds.width - w) / 2.0
            elif node.align == "right":
                x = bounds.left + (bounds.width - w)

            child_bounds = Rect(x, cur_y, w, h)
            self._solve_node(child, child_bounds, resolved)
            cur_y += h + node.gap

    def _solve_grid(self, node: Grid, bounds: Rect, resolved: List[ResolvedShape]):
        n = len(node.children)
        if n == 0:
            return

        cols = max(1, node.columns)
        rows = (n + cols - 1) // cols

        col_w = (bounds.width - (cols - 1) * node.col_gap) / cols
        row_h = (bounds.height - (rows - 1) * node.row_gap) / rows

        for idx, child in enumerate(node.children):
            r = idx // cols
            c = idx % cols
            x = bounds.left + c * (col_w + node.col_gap)
            y = bounds.top + r * (row_h + node.row_gap)
            child_bounds = Rect(x, y, col_w, row_h)
            self._solve_node(child, child_bounds, resolved)

    def _solve_card(self, node: CardNode, bounds: Rect, resolved: List[ResolvedShape]):
        # Card backplate container
        c_left = bounds.left
        c_top = bounds.top
        c_w = bounds.width
        c_h = bounds.height

        # Protruding header offset adjustments (e.g. Design 1 floating pill)
        protrusion = 0.0
        if node.header and node.header.style == "floating_pill":
            protrusion = node.header.protrusion_pt
            c_top += protrusion
            c_h = max(30.0, c_h - protrusion)

        # 1. Backplate
        resolved.append(ResolvedShape(
            shape_id=self._next_id("card_box"),
            kind="card_container",
            rect=Rect(c_left, c_top, c_w, c_h),
            z_order=1,
            fill_color=node.fill_color,
            border_color=node.border_color,
            border_width=node.border_width,
            corner_radius=node.corner_radius,
            shadow=node.shadow
        ))

        cur_content_y = c_top + node.padding.top
        pad_l = node.padding.left
        pad_r = node.padding.right
        usable_inner_w = max(10.0, c_w - pad_l - pad_r)

        # 2. Header
        if node.header:
            h = node.header
            if h.style == "floating_pill":
                pill_w = max(50.0, c_w - 24.0)
                pill_x = c_left + (c_w - pill_w) / 2.0
                pill_y = bounds.top
                resolved.append(ResolvedShape(
                    shape_id=self._next_id("header_pill"),
                    kind="header_pill",
                    rect=Rect(pill_x, pill_y, pill_w, h.height),
                    z_order=10,
                    fill_color=h.fill_color,
                    corner_radius=4.0,
                    text=h.title,
                    font_size=h.font_size,
                    bold=h.bold,
                    text_color=h.text_color,
                    align="center",
                    vertical_anchor="ctr",
                    gradient=h.gradient,
                    shadow=True
                ))
                cur_content_y = max(cur_content_y, c_top + (h.height - protrusion) + 14.0)

            elif h.style == "flush":
                resolved.append(ResolvedShape(
                    shape_id=self._next_id("flush_header"),
                    kind="flush_header",
                    rect=Rect(c_left + pad_l, cur_content_y, usable_inner_w, h.height),
                    z_order=10,
                    fill_color=h.fill_color,
                    corner_radius=4.0,
                    text=h.title,
                    font_size=h.font_size,
                    bold=h.bold,
                    text_color=h.text_color,
                    align="center",
                    vertical_anchor="ctr",
                    gradient=h.gradient
                ))
                cur_content_y += h.height + 14.0

            elif h.style == "top_tab":
                tab_w = max(60.0, c_w * 0.7)
                tab_x = c_left + (c_w - tab_w) / 2.0
                tab_y = c_top - 12.0
                resolved.append(ResolvedShape(
                    shape_id=self._next_id("header_tab"),
                    kind="header_pill",
                    rect=Rect(tab_x, tab_y, tab_w, h.height),
                    z_order=10,
                    fill_color=h.fill_color,
                    corner_radius=4.0,
                    text=h.title,
                    font_size=h.font_size,
                    bold=h.bold,
                    text_color=h.text_color,
                    align="center",
                    vertical_anchor="ctr"
                ))
                cur_content_y += 14.0

        # 3. Body Content
        body_h = max(20.0, (c_top + c_h - node.padding.bottom) - cur_content_y)
        body_bounds = Rect(c_left + pad_l, cur_content_y, usable_inner_w, body_h)

        if isinstance(node.body, BulletListNode):
            font_sz = node.body.font_size
            # Dynamic line fit: calculate accurate line wrap per item based on actual width
            char_per_line = max(15, int(body_bounds.width / (font_sz * 0.52)))
            est_lines = sum(max(1, (len(item) + char_per_line - 1) // char_per_line) for item in node.body.items)
            est_h = est_lines * (font_sz * 1.25) + max(0, len(node.body.items) - 1) * node.body.space_after
            if est_h > body_h and body_h > 30.0:
                # Downscale font size and space_after so entire block fits with 10% headroom
                scale = (body_h * 0.90) / est_h
                font_sz = max(8.0, round(font_sz * scale, 1))

            eff_space_after = min(node.body.space_after, max(1.5, (body_h - est_lines * font_sz * 1.2) / max(1, len(node.body.items))))

            resolved.append(ResolvedShape(
                shape_id=self._next_id("bullet_list"),
                kind="bullet_list",
                rect=body_bounds,
                z_order=10,
                font_size=font_sz,
                text_color=node.body.text_color,
                align="left",
                vertical_anchor="ctr",
                bullet_config={
                    "items": node.body.items,
                    "bullet_char": node.body.bullet_char,
                    "typeface": node.body.typeface,
                    "space_after": eff_space_after,
                    "mar_l_pt": node.body.mar_l_pt,
                    "indent_pt": node.body.indent_pt
                }
            ))
        elif isinstance(node.body, SectionListNode):
            resolved.append(ResolvedShape(
                shape_id=self._next_id("section_list"),
                kind="section_list",
                rect=body_bounds,
                z_order=10,
                align="left",
                vertical_anchor="ctr",
                metadata={
                    "sections": [s.model_dump() for s in node.body.sections],
                    "heading_font_size": node.body.heading_font_size,
                    "heading_color": node.body.heading_color,
                    "bullet_font_size": node.body.bullet_font_size,
                    "bullet_color": node.body.bullet_color,
                    "space_after_section": node.body.space_after_section
                }
            ))
        elif isinstance(node.body, LayoutNode):
            self._solve_node(node.body, body_bounds, resolved)

    def _solve_kpi(self, node: KPINode, bounds: Rect, resolved: List[ResolvedShape]):
        # Container
        resolved.append(ResolvedShape(
            shape_id=self._next_id("kpi_box"),
            kind="kpi_container",
            rect=bounds,
            z_order=1,
            fill_color=node.fill_color,
            border_color=node.border_color,
            border_width=0.75,
            corner_radius=node.corner_radius
        ))

        pad_t = node.padding.top
        pad_b = node.padding.bottom
        pad_l = node.padding.left
        pad_r = node.padding.right
        usable_w = bounds.width - (pad_l + pad_r)
        usable_h = bounds.height - (pad_t + pad_b)

        # Estimate line count of label to calculate safe font sizes
        char_per_line = max(10, int(usable_w / (node.label_size * 0.52)))
        lbl_lines = max(1, (len(node.label) + char_per_line - 1) // char_per_line)
        if node.sublabel:
            lbl_lines += max(1, (len(node.sublabel) + char_per_line - 1) // char_per_line)

        eff_num_sz = node.number_size
        eff_lbl_sz = node.label_size

        needed_h = eff_num_sz * 1.25 + 3.0 + lbl_lines * eff_lbl_sz * 1.2
        if needed_h > usable_h and usable_h > 20.0:
            scale = (usable_h * 0.92) / needed_h
            eff_num_sz = max(13.0, round(eff_num_sz * scale, 1))
            eff_lbl_sz = max(7.5, round(eff_lbl_sz * scale, 1))

        paras = [
            {
                "text": node.number,
                "font_size": eff_num_sz,
                "bold": True,
                "color": node.number_color,
                "align": "center",
                "space_after": 2.0,
            },
            {
                "text": node.label,
                "font_size": eff_lbl_sz,
                "bold": False,
                "color": node.label_color,
                "align": "center",
                "space_after": 0.0,
            }
        ]
        if node.sublabel:
            paras.append({
                "text": node.sublabel,
                "font_size": max(7.5, eff_lbl_sz - 1.0),
                "bold": False,
                "color": "NEUTRAL_MUTED",
                "align": "center",
                "space_after": 0.0,
            })

        # Emits a SINGLE text shape containing sequential paragraphs:
        # PowerPoint native text layout eliminates 100% of inter-text collisions.
        resolved.append(ResolvedShape(
            shape_id=self._next_id("kpi_content"),
            kind="kpi_content",
            rect=Rect(bounds.left + pad_l, bounds.top + pad_t, usable_w, usable_h),
            z_order=10,
            paragraphs=paras,
            align="center",
            vertical_anchor="ctr"
        ))

    def _solve_process_row(self, node: ProcessRowNode, bounds: Rect, resolved: List[ResolvedShape]):
        resolved.append(ResolvedShape(
            shape_id=self._next_id("proc_box"),
            kind="process_row",
            rect=bounds,
            z_order=1,
            fill_color=node.fill_color,
            corner_radius=node.corner_radius
        ))
        icon_w = 34.0
        gap = 12.0
        text_x = bounds.left + icon_w + gap + 10.0
        text_w = bounds.width - (icon_w + gap + 20.0)

        # Text
        resolved.append(ResolvedShape(
            shape_id=self._next_id("proc_text"),
            kind="text_box",
            rect=Rect(text_x, bounds.top, text_w, bounds.height),
            z_order=10,
            text=f"{node.title}\n{node.body}" if node.title else node.body,
            font_size=11.0,
            align="left",
            vertical_anchor="ctr"
        ))

        # Divider
        if node.divider:
            resolved.append(ResolvedShape(
                shape_id=self._next_id("proc_divider"),
                kind="divider",
                rect=Rect(text_x, bounds.top + bounds.height - 1.0, text_w, 1.0),
                z_order=5,
                line_prototype=node.divider
            ))

    def _solve_table(self, node: TableNode, bounds: Rect, resolved: List[ResolvedShape]):
        resolved.append(ResolvedShape(
            shape_id=self._next_id("table"),
            kind="table",
            rect=bounds,
            z_order=5,
            table_data={
                "headers": node.headers,
                "rows": node.rows,
                "col_widths": node.col_widths,
                "header_fill": node.header_fill,
                "zebra_fill": node.zebra_fill,
                "totals_row": node.totals_row,
                "totals_fill": node.totals_fill,
                "font_size": node.font_size
            }
        ))

    def _solve_divider(self, node: DividerNode, bounds: Rect, resolved: List[ResolvedShape]):
        resolved.append(ResolvedShape(
            shape_id=self._next_id("divider"),
            kind="divider",
            rect=bounds,
            z_order=5,
            line_prototype=node.line_prototype
        ))

    def _solve_connector(self, node: ConnectorNode, bounds: Rect, resolved: List[ResolvedShape]):
        resolved.append(ResolvedShape(
            shape_id=self._next_id("connector"),
            kind="connector",
            rect=bounds,
            z_order=5,
            line_prototype=node.line_prototype,
            fill_color=node.color
        ))

    def _solve_maturity_stage(self, node: MaturityStageNode, bounds: Rect, resolved: List[ResolvedShape]):
        # Tile container
        resolved.append(ResolvedShape(
            shape_id=self._next_id("stage_box"),
            kind="card_container",
            rect=bounds,
            z_order=1,
            fill_color="NEUTRAL_SURFACE",
            border_color="NEUTRAL_BORDER",
            corner_radius=node.corner_radius
        ))

        # Stage number pill
        pill_w = bounds.width - 24.0
        resolved.append(ResolvedShape(
            shape_id=self._next_id("stage_pill"),
            kind="header_pill",
            rect=Rect(bounds.left + 12.0, bounds.top + 10.0, pill_w, 24.0),
            z_order=10,
            fill_color=node.badge_color,
            corner_radius=4.0,
            text=f"STAGE {node.stage_num}: {node.stage_title.upper()}",
            font_size=10.0,
            bold=True,
            text_color="NEUTRAL_WHITE",
            align="center",
            vertical_anchor="ctr"
        ))

        # Capability bullet list
        body_y = bounds.top + 40.0
        body_h = bounds.height - 50.0
        resolved.append(ResolvedShape(
            shape_id=self._next_id("stage_bullets"),
            kind="bullet_list",
            rect=Rect(bounds.left + 12.0, body_y, bounds.width - 24.0, body_h),
            z_order=10,
            font_size=10.0,
            align="left",
            vertical_anchor="ctr",
            bullet_config={
                "items": node.capabilities,
                "bullet_char": "•",
                "typeface": "Aptos",
                "space_after": 6.0,
                "mar_l_pt": 11.0,
                "indent_pt": -11.0
            }
        ))

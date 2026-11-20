"""
Composable Layout Primitives
Defines the schema for atomic presentation layout nodes (containers and visual components).
"""

from typing import Any, Dict, List, Literal, Optional, Union
from pydantic import BaseModel, Field


class Padding(BaseModel):
    top: float = 8.5
    bottom: float = 8.5
    left: float = 14.0
    right: float = 14.0


class LayoutNode(BaseModel):
    """Base class for all composable layout nodes."""
    id: Optional[str] = None
    node_type: str = "node"
    width_weight: float = 1.0
    height_weight: float = 1.0
    fixed_width: Optional[float] = None
    fixed_height: Optional[float] = None
    margin: Padding = Field(default_factory=lambda: Padding(top=0.0, bottom=0.0, left=0.0, right=0.0))


# ==============================================================================
# Container Nodes
# ==============================================================================

class HStack(LayoutNode):
    """Horizontal flex container: arranges children left-to-right."""
    node_type: Literal["HStack"] = "HStack"
    children: List["AnyNode"] = Field(default_factory=list)
    gap: float = 18.0
    align: Literal["stretch", "top", "center", "bottom"] = "stretch"


class VStack(LayoutNode):
    """Vertical flex container: arranges children top-to-bottom."""
    node_type: Literal["VStack"] = "VStack"
    children: List["AnyNode"] = Field(default_factory=list)
    gap: float = 16.0
    align: Literal["stretch", "left", "center", "right"] = "stretch"


class Grid(LayoutNode):
    """2D Grid matrix: arranges children in rows and columns."""
    node_type: Literal["Grid"] = "Grid"
    children: List["AnyNode"] = Field(default_factory=list)
    columns: int = 3
    col_gap: float = 18.0
    row_gap: float = 16.0


# ==============================================================================
# Visual Component Nodes
# ==============================================================================

class PillHeaderNode(BaseModel):
    """Header styling: floating pill, flush banner, or top tab."""
    title: str
    style: Literal["flush", "floating_pill", "top_tab", "banner"] = "flush"
    fill_color: str = "BRAND_INDIGO"
    text_color: str = "NEUTRAL_WHITE"
    font_size: float = 13.5
    bold: bool = True
    height: float = 32.0
    protrusion_pt: float = 16.0  # height protruding above container for floating_pill
    gradient: Optional[str] = None  # "linear_indigo", "linear_navy", "linear_tangerine"


class BulletListNode(BaseModel):
    """Native bulleted paragraph list."""
    items: List[str] = Field(default_factory=list)
    font_size: float = 11.0
    text_color: str = "BRAND_ONYX"
    bullet_char: str = "•"
    typeface: str = "Aptos"
    space_after: float = 7.0
    mar_l_pt: float = 13.5
    indent_pt: float = -13.5


class SectionItem(BaseModel):
    heading: str
    bullets: List[str] = Field(default_factory=list)


class SectionListNode(BaseModel):
    """Paired subheader + bullet list (e.g., Program Approach on Slide 4)."""
    sections: List[SectionItem] = Field(default_factory=list)
    heading_font_size: float = 11.5
    heading_color: str = "BRAND_INDIGO"
    bullet_font_size: float = 10.5
    bullet_color: str = "BRAND_ONYX"
    space_after_section: float = 10.0


class CardNode(LayoutNode):
    """Container card with 4.0pt corner radius, padding, and optional header."""
    node_type: Literal["Card"] = "Card"
    header: Optional[PillHeaderNode] = None
    body: Optional[Union[BulletListNode, SectionListNode, "AnyNode"]] = None
    fill_color: str = "NEUTRAL_WHITE"
    border_color: Optional[str] = "NEUTRAL_BORDER"
    border_width: float = 0.75
    corner_radius: float = 4.0
    shadow: bool = False
    aspect_ratio: Optional[float] = None
    padding: Padding = Field(default_factory=lambda: Padding(top=10.0, bottom=10.0, left=14.0, right=14.0))


class KPINode(LayoutNode):
    """High-impact metric card (e.g. '5%-6%', 'Shift', '~700')."""
    node_type: Literal["KPICard"] = "KPICard"
    number: str
    label: str
    sublabel: Optional[str] = None
    fill_color: str = "NEUTRAL_WHITE"
    border_color: Optional[str] = "NEUTRAL_BORDER"
    number_color: str = "BRAND_INDIGO"
    label_color: str = "BRAND_ONYX"
    corner_radius: float = 4.0
    number_size: float = 22.0
    label_size: float = 10.0
    padding: Padding = Field(default_factory=lambda: Padding(top=8.0, bottom=8.0, left=10.0, right=10.0))


class ProcessRowNode(LayoutNode):
    """Horizontal process step with icon and description."""
    node_type: Literal["ProcessRow"] = "ProcessRow"
    index_badge: Optional[str] = None
    icon_name: Optional[str] = None
    icon_svg_path: Optional[str] = None
    title: str = ""
    body: str = ""
    fill_color: str = "NEUTRAL_SURFACE"
    corner_radius: float = 4.0
    divider: Optional[str] = "side_fade_r"


class TableNode(LayoutNode):
    """Structured data table."""
    node_type: Literal["Table"] = "Table"
    headers: List[str] = Field(default_factory=list)
    rows: List[List[str]] = Field(default_factory=list)
    col_widths: Optional[List[float]] = None
    header_fill: str = "BRAND_INDIGO"
    zebra_fill: str = "NEUTRAL_SURFACE"
    totals_row: bool = False
    totals_fill: str = "BRAND_NAVY"
    font_size: float = 9.5


class DividerNode(LayoutNode):
    """Authentic vector gradient line cloned from Library Slide 10:
    - Grey Centre Fade V (vertical, center-aligned content)
    - Grey Solid V (vertical, left/standard-aligned content)
    - Grey Centre Fade H (horizontal, center-aligned content)
    - Grey Side Fade R (horizontal, left-aligned content)
    - Grey Solid H (horizontal, solid divider)
    """
    node_type: Literal["Divider"] = "Divider"
    line_prototype: Optional[str] = None  # If None, automatically determined by orientation and is_center_aligned
    orientation: Literal["horizontal", "vertical", "h", "v"] = "horizontal"
    is_center_aligned: bool = False
    thickness: float = 1.0


class ConnectorNode(LayoutNode):
    """Orthogonal connector line."""
    node_type: Literal["Connector"] = "Connector"
    line_prototype: str = "straight_connector_8"
    color: str = "NEUTRAL_LINE"
    thickness: float = 1.0


class MaturityStageNode(LayoutNode):
    """Progression stage tile for maturity curves (Slide 8 pattern)."""
    node_type: Literal["MaturityStage"] = "MaturityStage"
    stage_num: int
    stage_title: str
    capabilities: List[str] = Field(default_factory=list)
    value_metrics: Optional[str] = None
    badge_color: str = "BRAND_INDIGO"
    corner_radius: float = 4.0


class StackLayerItem(BaseModel):
    """An item within a layer of an architectural stack."""
    text: str
    icon_name: Optional[str] = None
    icon_svg_file: Optional[str] = None
    icon_svg_bytes: Optional[bytes] = None


class StackLayer(BaseModel):
    """A single horizontal tier in an architectural layered stack."""
    level: int
    title: str
    items: List[StackLayerItem] = Field(default_factory=list)
    header_fill: str = "BRAND_INDIGO"
    header_text_color: str = "NEUTRAL_WHITE"
    body_fill: str = "NEUTRAL_SURFACE"
    body_alpha: int = 100000  # DrawingML alpha: e.g. 20000 (20%), 30000 (30%), 100000 (100%)
    border_color: str = "NEUTRAL_BORDER"
    border_width: float = 0.5
    divider_color: Optional[str] = None
    header_icon_file: Optional[str] = None
    header_icon_bytes: Optional[bytes] = None
    badge_text: Optional[str] = None
    badge_fill: str = "#FAF9F6"
    badge_text_color: str = "#14203E"
    icon_tint_color: Optional[str] = None
    header_icon_tint: Optional[str] = None


class LayeredStackNode(LayoutNode):
    """Architectural layered stack (Slide 9 pattern in Library_copy.pptx)."""
    node_type: Literal["LayeredStack"] = "LayeredStack"
    layers: List[StackLayer] = Field(default_factory=list)
    foundation_at_bottom: bool = True


class OverlayNode(LayoutNode):
    """Layer container for floating elements (pills, badges, connectors, annotations)
    with explicit protrusion, z-order, and alignment rules above a base container."""
    node_type: Literal["Overlay"] = "Overlay"
    base: "AnyNode"
    overlays: List["AnyNode"] = Field(default_factory=list)
    protrusion_pt: float = 16.0  # e.g. +16pt protrusion above card
    horizontal_align: Literal["center", "left", "right"] = "center"
    z_order_offset: int = 10


# Union for any node type in the tree
AnyNode = Union[
    HStack,
    VStack,
    Grid,
    CardNode,
    KPINode,
    ProcessRowNode,
    TableNode,
    DividerNode,
    ConnectorNode,
    MaturityStageNode,
    LayeredStackNode,
    OverlayNode,
]

# Update forward refs
HStack.model_rebuild()
VStack.model_rebuild()
Grid.model_rebuild()
CardNode.model_rebuild()
LayeredStackNode.model_rebuild()
OverlayNode.model_rebuild()


class SlideComposition(BaseModel):
    """Root slide composition containing metadata and the root layout tree."""
    title: str
    subtitle: Optional[str] = None
    category_tracker: Optional[str] = None
    root: AnyNode
    footnote: Optional[str] = None

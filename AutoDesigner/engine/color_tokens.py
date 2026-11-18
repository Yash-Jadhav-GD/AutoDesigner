"""
AutoDesigner Color Tokens & Color Theory Engine
Strict DrawingML Schema-compliant fill and effect injection:
DrawingML CT_ShapeProperties sequence:
  1. xfrm
  2. prstGeom / custGeom
  3. [fill] (noFill | solidFill | gradFill | blipFill | pattFill | grpFill)
  4. ln
  5. effectLst
  6. scene3d / sp3d / extLst
"""

from typing import Tuple, Dict, Any, Optional
from pptx.oxml import parse_xml
from pptx.dml.color import RGBColor


class ColorTokens:
    # 1. Core Brand Palette (Layer 1: Canonical)
    BRAND_INDIGO = "48397E"       # 72, 57, 126
    BRAND_TANGERINE = "FE860E"    # 254, 134, 14
    BRAND_NAVY = "1F2E65"         # 31, 46, 101
    BRAND_ONYX = "221E1F"         # 34, 30, 31

    # 2. Approved Vivid Tints (Layer 2: Benchmark)
    VIVID_INDIGO_MID = "7C6DB2"
    VIVID_INDIGO_SOFT = "B1A8D8"
    VIVID_INDIGO_WASH = "E2DEEE"
    VIVID_INDIGO_LIGHT = "C1BAE0"

    VIVID_NAVY_MID = "526299"
    VIVID_NAVY_SOFT = "8F9BC7"

    VIVID_TANGERINE_MID = "FFAC59"
    VIVID_TANGERINE_SOFT = "FFD3A6"
    VIVID_TANGERINE_WASH = "FFECD9"

    # 3. Neutrals
    NEUTRAL_WHITE = "FFFFFF"
    NEUTRAL_SURFACE = "F6F8FB"
    NEUTRAL_BORDER = "DFE3E6"
    NEUTRAL_LINE = "A2ABB0"

    # 4. Text Tokens
    TEXT_PRIMARY = "262626"
    TEXT_HEADING = "221E1F"
    TEXT_SECONDARY = "475569"
    TEXT_ON_DARK = "FFFFFF"
    TEXT_ACCENT = "48397E"

    # 5. Slide 1 Linear Gradient Pill Configurations (After_copy.pptx Benchmark)
    GRADIENT_INDIGO_START = "6A5E96"
    GRADIENT_INDIGO_END = "48397E"

    GRADIENT_NAVY_START = "5A668E"
    GRADIENT_NAVY_END = "1F2E65"

    GRADIENT_TANGERINE_START = "FEB163"
    GRADIENT_TANGERINE_END = "FE860E"

    # Shadow settings (authentic from Decor_Box_with_Shadow_1 on Library Slide 6)
    SHADOW_COLOR = "95939D"
    SHADOW_ALPHA = 16000          # 16%
    SHADOW_BLUR_RAD = 127000      # 10 pt blur (127000 EMU)

    # 5-Phase Process Ramp (After_copy.pptx Slide 6 & 20)
    PHASE_RAMP = [
        VIVID_TANGERINE_MID,   # Phase 1: FFAC59
        VIVID_TANGERINE_SOFT,  # Phase 2: FFD3A6
        VIVID_INDIGO_MID,      # Phase 3: 7C6DB2
        VIVID_INDIGO_SOFT,     # Phase 4: B1A8D8
        VIVID_NAVY_SOFT,       # Phase 5: 8F9BC7
    ]


class TypographyTokens:
    """Typography standards from AUTOFORMAT_SPECIFICATION.md section 7 (sizes in pt).
    Font, colour and alignment come from the cloned Library prototype; these set the sizes."""
    FONT_FAMILY = "Aptos"
    KPI_NUMBER = 28.0              # Bold, centred
    KPI_HEADER = 14.0              # SemiBold
    PILL_HEADER = 16.0             # Bold, #FFFFFF
    FLUSH_HEADER = 16.0            # Bold, #FFFFFF (14-16)
    SUBHEADER = 14.0               # SemiBold
    SUBHEADER_COLOR = "2D3E50"
    BODY = 11.0                    # Regular, locked across card designs
    BODY_MIN = 9.0                 # last resort when a slide's text cannot fit at BODY (text is never dropped)
    BODY_COLOR = "262626"
    KPI_BODY = 10.0                # Regular, centred
    KPI_BODY_COLOR = "151514"
    CARD_PADDING = 8.5039          # 0.3 cm inner padding, top and bottom
    MIN_GAP = 5.0                  # section 5: minimum gap above and below a divider
    ICON_MIN = 15.0                # section 6: uniform icon size floor

    # Text-height estimates (python-pptx cannot measure text)
    LINE_SPACING = 1.2             # line height / font size
    CHAR_WIDTH_EM = 0.46           # average Aptos character width / font size
    CHAR_WIDTH_EM_BOLD = 0.51


def brand_palette() -> Dict[str, str]:
    """Palette-shaped view of ColorTokens, matching the keys in templates/palettes.json."""
    return {
        "name": "GEP Brand (Indigo / Tangerine)",
        "primary": ColorTokens.BRAND_INDIGO,
        "secondary": ColorTokens.BRAND_NAVY,
        "accent": ColorTokens.BRAND_TANGERINE,
        "card_bg": ColorTokens.NEUTRAL_WHITE,
        "card_border": ColorTokens.NEUTRAL_BORDER,
        "slide_bg": ColorTokens.NEUTRAL_SURFACE,
        "text_dark": ColorTokens.TEXT_PRIMARY,
        "text_muted": ColorTokens.TEXT_SECONDARY,
        "text_light": ColorTokens.TEXT_ON_DARK,
        "divider": ColorTokens.NEUTRAL_LINE,
        "badge_bg": ColorTokens.VIVID_INDIGO_WASH,
    }


def hex_to_rgb(hex_code: str) -> RGBColor:
    """Converts a 6-character hex string to pptx.dml.color.RGBColor."""
    clean = hex_code.lstrip("#")
    return RGBColor(int(clean[0:2], 16), int(clean[2:4], 16), int(clean[4:6], 16))


def clear_fill_and_effects(spPr):
    """Removes existing fill and effect elements from spPr."""
    tags_to_remove = [
        "{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill",
        "{http://schemas.openxmlformats.org/drawingml/2006/main}gradFill",
        "{http://schemas.openxmlformats.org/drawingml/2006/main}noFill",
        "{http://schemas.openxmlformats.org/drawingml/2006/main}blipFill",
        "{http://schemas.openxmlformats.org/drawingml/2006/main}pattFill",
        "{http://schemas.openxmlformats.org/drawingml/2006/main}grpFill",
        "{http://schemas.openxmlformats.org/drawingml/2006/main}effectLst",
    ]
    for child in list(spPr):
        if child.tag in tags_to_remove:
            spPr.remove(child)


def _insert_fill_schema_compliant(spPr, fill_elem):
    """
    Inserts a fill element into spPr strictly adhering to DrawingML schema:
    Fill MUST appear before <a:ln>, <a:effectLst>, <a:scene3d>, <a:sp3d>, <a:extLst>.
    """
    ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    ln = spPr.find(f"{{{ns_a}}}ln")
    effectLst = spPr.find(f"{{{ns_a}}}effectLst")
    scene3d = spPr.find(f"{{{ns_a}}}scene3d")
    sp3d = spPr.find(f"{{{ns_a}}}sp3d")
    extLst = spPr.find(f"{{{ns_a}}}extLst")

    for successor in [ln, effectLst, scene3d, sp3d, extLst]:
        if successor is not None:
            successor.addprevious(fill_elem)
            return

    spPr.append(fill_elem)


def _insert_effect_schema_compliant(spPr, effect_elem):
    """
    Inserts an effectLst element into spPr strictly adhering to DrawingML schema:
    effectLst MUST appear after <a:ln> and before <a:scene3d>, <a:sp3d>, <a:extLst>.
    """
    ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    scene3d = spPr.find(f"{{{ns_a}}}scene3d")
    sp3d = spPr.find(f"{{{ns_a}}}sp3d")
    extLst = spPr.find(f"{{{ns_a}}}extLst")

    for successor in [scene3d, sp3d, extLst]:
        if successor is not None:
            successor.addprevious(effect_elem)
            return

    spPr.append(effect_elem)


def apply_linear_gradient(
    shape,
    start_hex: str,
    end_hex: str,
    angle_dml: int = 10800000,
    with_shadow: bool = True
):
    """
    Applies a vertical/directional 2-stop linear gradient and drop shadow to a shape.
    angle_dml 10800000 = 90 degrees (top to bottom).
    Ensures strict DrawingML schema sequence: (xfrm, prstGeom, gradFill, ln, effectLst).
    """
    spPr = shape._element.spPr
    clear_fill_and_effects(spPr)

    # 1. Gradient Fill (inserted BEFORE <a:ln>)
    grad_xml = f"""
    <a:gradFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" flip="none" rotWithShape="1">
      <a:gsLst>
        <a:gs pos="0">
          <a:srgbClr val="{start_hex.lstrip('#')}"/>
        </a:gs>
        <a:gs pos="100000">
          <a:srgbClr val="{end_hex.lstrip('#')}"/>
        </a:gs>
      </a:gsLst>
      <a:lin ang="{angle_dml}" scaled="1"/>
      <a:tileRect/>
    </a:gradFill>
    """
    _insert_fill_schema_compliant(spPr, parse_xml(grad_xml))

    # 2. Drop Shadow (inserted AFTER <a:ln>, only if shape does not contain opacity)
    if with_shadow:
        # Check if shape contains opacity: if alpha < 100000, do not apply shadow
        alpha_elem = spPr.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}alpha")
        if alpha_elem is not None and int(alpha_elem.get("val", "100000")) < 100000:
            return
        shdw_xml = f"""
        <a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:outerShdw blurRad="{ColorTokens.SHADOW_BLUR_RAD}" rotWithShape="0">
            <a:srgbClr val="{ColorTokens.SHADOW_COLOR}">
              <a:alpha val="{ColorTokens.SHADOW_ALPHA}"/>
            </a:srgbClr>
          </a:outerShdw>
        </a:effectLst>
        """
        _insert_effect_schema_compliant(spPr, parse_xml(shdw_xml))


def apply_solid_fill(shape, hex_val: str):
    """
    Applies a solid fill cleanly to a shape via OpenXML,
    strictly inserting before <a:ln> to comply with DrawingML schema.
    """
    spPr = shape._element.spPr
    clear_fill_and_effects(spPr)

    fill_xml = f"""
    <a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:srgbClr val="{hex_val.lstrip('#')}"/>
    </a:solidFill>
    """
    _insert_fill_schema_compliant(spPr, parse_xml(fill_xml))

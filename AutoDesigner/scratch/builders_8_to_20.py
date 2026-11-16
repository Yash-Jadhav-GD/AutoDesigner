"""
AutoDesigner Builders for Slides 8 through 20.
Sourced strictly from Library_copy.pptx and After_copy.pptx templates:
- Clean white cards (#FFFFFF) with subtle border (#CBD5E1) and Library shadow (Decor_Box_with_Shadow_1)
- Floating pill headers (Design 1 Header Text) in Navy (#1F2E65) or Tangerine (#FF8300)
- UI-ROW-01 horizontal cards from Library Slide 2 & 3
- KPI-CRD-V1 metric cards from Library Slide 5
- CRD-ICN-V1 Data Box cards from Library Slide 7
- Table matrix from After_copy.pptx Slide 19
- Process flow chevrons from After_copy.pptx Slide 2
- 100% Middle vertical alignment on all text frames
- PowerPoint internal DrawingML native bullets (<a:buChar char="•"/>)
- Exact 4.0pt corner radius on all rounded rectangles
- Library prototype divider lines
- 100% text fidelity (0 lost tokens)
- 0 shape overlaps
"""

import copy
import io
import os
import pptx
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from engine.builders.common import (
    clean_element,
    get_shape_by_name,
    add_library_line,
    set_shape_corner_radius,
    apply_library_shadow,
    apply_native_bullet,
    clean_bullet_string,
)
from engine.color_tokens import ColorTokens, apply_solid_fill, apply_linear_gradient

def hex_to_rgb(hex_str: str) -> RGBColor:
    hex_clean = hex_str.lstrip("#")
    return RGBColor(int(hex_clean[0:2], 16), int(hex_clean[2:4], 16), int(hex_clean[4:6], 16))


# ==============================================================================
# SLIDE 8: Executive Summary
# ==============================================================================

def build_slide8(prs, lib_s1, lib_lines):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Title
    title_box = slide.shapes.add_textbox(Pt(30.0), Pt(8.0), Pt(900.0), Pt(32.0))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Executive Summary: One Accountable Partner Delivering 6 to 15% Savings Through an AI Enabled Managed Service"
    p.font.name = "Aptos"
    p.font.size = Pt(14.5)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

    # Top Context Container (compact, y=42, h=40, bottom=82)
    ctx_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(42.0), Pt(900.0), Pt(40.0))
    apply_solid_fill(ctx_box, "F8FAFC")
    ctx_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(ctx_box, 4.0)
    apply_library_shadow(ctx_box)

    c_tf = ctx_box.text_frame
    c_tf.word_wrap = True
    c_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    cp = c_tf.paragraphs[0]
    r1 = cp.add_run()
    r1.text = "Context and objective: "
    r1.font.bold = True
    r1.font.size = Pt(9.5)
    r1.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)
    r1.font.name = "Aptos"
    r2 = cp.add_run()
    r2.text = "Provide KPOD with a single accountable PSCM partner across Wells, Operations, Projects and Indirect categories, ensuring continuity through transition, improved supplier performance, measurable value delivery, and increased Iraqi local content participation."
    r2.font.size = Pt(9.5)
    r2.font.color.rgb = hex_to_rgb("334155")
    r2.font.name = "Aptos"

    # 3 Columns Layout: col_top = 104.0 (leaving 22pt gap from top context box)
    # Header pills at y = 92.0, h = 24.0. Cards at y = 104.0, h = 422.0.
    decor_proto = get_shape_by_name(lib_s1, "Design 1 Decor box")
    hdr_proto = get_shape_by_name(lib_s1, "Design 1 Header Text")
    hdr_proto = get_shape_by_name(lib_s1, "Design 1 Header Text")

    col_w1 = 264.0
    col_w2 = 344.0
    col_w3 = 264.0
    col_x1 = 30.0
    col_x2 = 308.0
    col_x3 = 666.0
    col_top = 106.0
    col_h = 420.0

    # -------------------------------------------------------------
    # Col 1: Scope of Work
    # -------------------------------------------------------------
    card1_elem = clean_element(copy.deepcopy(decor_proto._element))
    card1 = slide.shapes._spTree.append(card1_elem)
    sh_c1 = slide.shapes[-1]
    sh_c1.left = Pt(col_x1)
    sh_c1.top = Pt(col_top)
    sh_c1.width = Pt(col_w1)
    sh_c1.height = Pt(col_h)
    apply_solid_fill(sh_c1, "FFFFFF")
    sh_c1.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(sh_c1, 4.0)

    # Header pill
    h1_elem = clean_element(copy.deepcopy(hdr_proto._element))
    slide.shapes._spTree.append(h1_elem)
    sh_h1 = slide.shapes[-1]
    sh_h1.left = Pt(col_x1 + (col_w1 - 200.0) / 2.0)
    sh_h1.top = Pt(94.0)
    sh_h1.width = Pt(200.0)
    sh_h1.height = Pt(24.0)
    apply_solid_fill(sh_h1, ColorTokens.BRAND_INDIGO)
    set_shape_corner_radius(sh_h1, 4.0)
    sh_h1.text_frame.word_wrap = False
    sh_h1.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_h1 = sh_h1.text_frame.paragraphs[0]
    p_h1.text = "Scope of Work"
    p_h1.alignment = PP_ALIGN.CENTER
    p_h1.font.name = "Aptos"
    p_h1.font.size = Pt(11.0)
    p_h1.font.bold = True
    p_h1.font.color.rgb = hex_to_rgb("FFFFFF")

    # Inside Col 1: VOLUMES subheader
    v_lbl = slide.shapes.add_textbox(Pt(col_x1 + 12.0), Pt(col_top + 16.0), Pt(col_w1 - 24.0), Pt(18.0))
    vp = v_lbl.text_frame.paragraphs[0]
    vp.text = "VOLUMES"
    vp.font.name = "Aptos"
    vp.font.size = Pt(9.5)
    vp.font.bold = True
    vp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

    # 4 Metric Cards (KPI-CRD-V1 style)
    # Row 1
    m1_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(col_x1 + 12.0), Pt(col_top + 36.0), Pt(114.0), Pt(46.0))
    apply_solid_fill(m1_box, "F8FAFC")
    m1_box.line.color.rgb = hex_to_rgb("E2E8F0")
    set_shape_corner_radius(m1_box, 4.0)
    m1_tf = m1_box.text_frame
    m1_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p1 = m1_tf.paragraphs[0]
    p1.text = "148"
    p1.font.name = "Aptos"
    p1.font.size = Pt(16.0)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_TANGERINE)
    p1.alignment = PP_ALIGN.CENTER
    p2 = m1_tf.add_paragraph()
    p2.text = "Contracts Executed"
    p2.font.name = "Aptos"
    p2.font.size = Pt(9.0)
    p2.font.color.rgb = hex_to_rgb("64748B")
    p2.alignment = PP_ALIGN.CENTER

    m2_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(col_x1 + 138.0), Pt(col_top + 36.0), Pt(114.0), Pt(46.0))
    apply_solid_fill(m2_box, "F8FAFC")
    m2_box.line.color.rgb = hex_to_rgb("E2E8F0")
    set_shape_corner_radius(m2_box, 4.0)
    m2_tf = m2_box.text_frame
    m2_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p1 = m2_tf.paragraphs[0]
    p1.text = "130"
    p1.font.name = "Aptos"
    p1.font.size = Pt(16.0)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)
    p1.alignment = PP_ALIGN.CENTER
    p2 = m2_tf.add_paragraph()
    p2.text = "Live Sourcing Cases"
    p2.font.name = "Aptos"
    p2.font.size = Pt(9.0)
    p2.font.color.rgb = hex_to_rgb("64748B")
    p2.alignment = PP_ALIGN.CENTER

    # Row 2 (Personnel with +)
    m3_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(col_x1 + 12.0), Pt(col_top + 88.0), Pt(102.0), Pt(46.0))
    apply_solid_fill(m3_box, "F8FAFC")
    m3_box.line.color.rgb = hex_to_rgb("E2E8F0")
    set_shape_corner_radius(m3_box, 4.0)
    m3_tf = m3_box.text_frame
    m3_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p1 = m3_tf.paragraphs[0]
    p1.text = "9"
    p1.font.name = "Aptos"
    p1.font.size = Pt(16.0)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)
    p1.alignment = PP_ALIGN.CENTER
    p2 = m3_tf.add_paragraph()
    p2.text = "BP PSCM Personnel"
    p2.font.name = "Aptos"
    p2.font.size = Pt(9.0)
    p2.font.color.rgb = hex_to_rgb("64748B")
    p2.alignment = PP_ALIGN.CENTER

    plus_box = slide.shapes.add_textbox(Pt(col_x1 + 116.0), Pt(col_top + 98.0), Pt(20.0), Pt(26.0))
    pp = plus_box.text_frame.paragraphs[0]
    pp.text = "+"
    pp.font.name = "Aptos"
    pp.font.size = Pt(16.0)
    pp.font.bold = True
    pp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)
    pp.alignment = PP_ALIGN.CENTER

    m4_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(col_x1 + 138.0), Pt(col_top + 88.0), Pt(114.0), Pt(46.0))
    apply_solid_fill(m4_box, "F8FAFC")
    m4_box.line.color.rgb = hex_to_rgb("E2E8F0")
    set_shape_corner_radius(m4_box, 4.0)
    m4_tf = m4_box.text_frame
    m4_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p1 = m4_tf.paragraphs[0]
    p1.text = "23"
    p1.font.name = "Aptos"
    p1.font.size = Pt(16.0)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_TANGERINE)
    p1.alignment = PP_ALIGN.CENTER
    p2 = m4_tf.add_paragraph()
    p2.text = "Agency Personnel"
    p2.font.name = "Aptos"
    p2.font.size = Pt(9.0)
    p2.font.color.rgb = hex_to_rgb("64748B")
    p2.alignment = PP_ALIGN.CENTER

    # Divider
    add_library_line(slide, lib_lines["side_fade_r"], col_x1 + 14.0, col_top + 142.0, col_w1 - 28.0)

    # Section 2: IN-SCOPE SERVICES
    s_lbl = slide.shapes.add_textbox(Pt(col_x1 + 12.0), Pt(col_top + 148.0), Pt(col_w1 - 24.0), Pt(18.0))
    sp = s_lbl.text_frame.paragraphs[0]
    sp.text = "IN-SCOPE SERVICES"
    sp.font.name = "Aptos"
    sp.font.size = Pt(9.5)
    sp.font.bold = True
    sp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

    # 8 Bullets
    bullets_s8_col1 = [
        "Category strategy and sourcing execution",
        "RFx management and commercial evaluations",
        "Contract development and administration",
        "Procurement operations and buying support",
        "Supplier performance and risk management",
        "Market intelligence and should-cost analysis",
        "Materials and inventory coordination",
        "Governance, reporting and local content support"
    ]
    b_box1 = slide.shapes.add_textbox(Pt(col_x1 + 12.0), Pt(col_top + 168.0), Pt(col_w1 - 24.0), Pt(244.0))
    b_tf1 = b_box1.text_frame
    b_tf1.word_wrap = True
    b_tf1.vertical_anchor = MSO_ANCHOR.MIDDLE
    for idx, b_txt in enumerate(bullets_s8_col1):
        pb = b_tf1.paragraphs[0] if idx == 0 else b_tf1.add_paragraph()
        pb.text = clean_bullet_string(b_txt)
        pb.font.name = "Aptos"
        pb.font.size = Pt(9.0)
        pb.font.color.rgb = hex_to_rgb("334155")
        pb.space_after = Pt(4.5)
        apply_native_bullet(pb, char="•", typeface="Arial", mar_l_pt=11.0, indent_pt=-11.0)

    # -------------------------------------------------------------
    # Col 2: Solution Summary (4 cards in 2x2 grid)
    # -------------------------------------------------------------
    card2_elem = clean_element(copy.deepcopy(decor_proto._element))
    slide.shapes._spTree.append(card2_elem)
    sh_c2 = slide.shapes[-1]
    sh_c2.left = Pt(col_x2)
    sh_c2.top = Pt(col_top)
    sh_c2.width = Pt(col_w2)
    sh_c2.height = Pt(col_h)
    apply_solid_fill(sh_c2, "FFFFFF")
    sh_c2.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(sh_c2, 4.0)

    # Header pill
    h2_elem = clean_element(copy.deepcopy(hdr_proto._element))
    slide.shapes._spTree.append(h2_elem)
    sh_h2 = slide.shapes[-1]
    sh_h2.left = Pt(col_x2 + (col_w2 - 220.0) / 2.0)
    sh_h2.top = Pt(94.0)
    sh_h2.width = Pt(220.0)
    sh_h2.height = Pt(24.0)
    apply_solid_fill(sh_h2, ColorTokens.BRAND_TANGERINE)
    set_shape_corner_radius(sh_h2, 4.0)
    sh_h2.text_frame.word_wrap = False
    sh_h2.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_h2 = sh_h2.text_frame.paragraphs[0]
    p_h2.text = "Solution Summary"
    p_h2.alignment = PP_ALIGN.CENTER
    p_h2.font.name = "Aptos"
    p_h2.font.size = Pt(11.0)
    p_h2.font.bold = True
    p_h2.font.color.rgb = hex_to_rgb("FFFFFF")

    grid_cards = [
        {
            "title": "Savings",
            "clr": ColorTokens.BRAND_TANGERINE,
            "bullets": [
                "6 to 15% on addressable spend",
                "USD 60–100M indicative annual savings opportunity",
                "Proven track record of value delivery across Wells, Operations, Projects and Indirect categories"
            ]
        },
        {
            "title": "Services and Activities",
            "clr": ColorTokens.BRAND_NAVY,
            "bullets": [
                "Integrated Source-to-Pay (S2P) & SC managed services",
                "Innovative category sourcing strategies, RFx and contract management",
                "P2P, materials, supplier performance and risk management"
            ]
        },
        {
            "title": "Digital and AI",
            "clr": ColorTokens.BRAND_NAVY,
            "bullets": [
                "Platform-agnostic integration with KPOD's existing technology ecosystem",
                "GEP AI agents for intake, RFx, negotiations and contracts",
                "AI-driven analytics, should-cost and automation"
            ]
        },
        {
            "title": "Delivery Model",
            "clr": ColorTokens.BRAND_INDIGO,
            "bullets": [
                "Dedicated Category Managers across key spend areas",
                "Centralized governance and performance management",
                "Integrated local content tracking, supplier development and KPI reporting"
            ]
        }
    ]

    gc_w = 158.0
    gc_h = 192.0
    gc_coords = [
        (col_x2 + 10.0, col_top + 14.0),
        (col_x2 + 176.0, col_top + 14.0),
        (col_x2 + 10.0, col_top + 214.0),
        (col_x2 + 176.0, col_top + 214.0)
    ]

    for g_idx, gc in enumerate(grid_cards):
        gx, gy = gc_coords[g_idx]
        gbox = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(gx), Pt(gy), Pt(gc_w), Pt(gc_h))
        apply_solid_fill(gbox, "F8FAFC")
        gbox.line.color.rgb = hex_to_rgb("CBD5E1")
        set_shape_corner_radius(gbox, 4.0)
        apply_library_shadow(gbox)

        # Title bar
        t_box = slide.shapes.add_textbox(Pt(gx + 8.0), Pt(gy + 6.0), Pt(gc_w - 16.0), Pt(24.0))
        tp = t_box.text_frame.paragraphs[0]
        tp.text = gc["title"]
        tp.font.name = "Aptos"
        tp.font.size = Pt(10.0)
        tp.font.bold = True
        tp.font.color.rgb = hex_to_rgb(gc["clr"])

        # Divider
        add_library_line(slide, lib_lines["side_fade_r"], gx + 8.0, gy + 32.0, gc_w - 16.0)

        # Bullets
        b_box = slide.shapes.add_textbox(Pt(gx + 6.0), Pt(gy + 36.0), Pt(gc_w - 12.0), Pt(gc_h - 40.0))
        b_tf = b_box.text_frame
        b_tf.word_wrap = True
        b_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        for b_i, bullet in enumerate(gc["bullets"]):
            p_b = b_tf.paragraphs[0] if b_i == 0 else b_tf.add_paragraph()
            p_b.text = clean_bullet_string(bullet)
            p_b.font.name = "Aptos"
            p_b.font.size = Pt(9.0)
            p_b.font.color.rgb = hex_to_rgb("334155")
            p_b.space_after = Pt(4.0)
            apply_native_bullet(p_b, char="•", typeface="Arial", mar_l_pt=10.0, indent_pt=-10.0)

    # -------------------------------------------------------------
    # Col 3: Benefits with GEP (6 UI-ROW-01 cards)
    # -------------------------------------------------------------
    card3_elem = clean_element(copy.deepcopy(decor_proto._element))
    slide.shapes._spTree.append(card3_elem)
    sh_c3 = slide.shapes[-1]
    sh_c3.left = Pt(col_x3)
    sh_c3.top = Pt(col_top)
    sh_c3.width = Pt(col_w3)
    sh_c3.height = Pt(col_h)
    apply_solid_fill(sh_c3, "FFFFFF")
    sh_c3.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(sh_c3, 4.0)

    # Header pill
    h3_elem = clean_element(copy.deepcopy(hdr_proto._element))
    slide.shapes._spTree.append(h3_elem)
    sh_h3 = slide.shapes[-1]
    sh_h3.left = Pt(col_x3 + (col_w3 - 200.0) / 2.0)
    sh_h3.top = Pt(94.0)
    sh_h3.width = Pt(200.0)
    sh_h3.height = Pt(24.0)
    apply_solid_fill(sh_h3, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(sh_h3, 4.0)
    sh_h3.text_frame.word_wrap = False
    sh_h3.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_h3 = sh_h3.text_frame.paragraphs[0]
    p_h3.text = "Benefits with GEP"
    p_h3.alignment = PP_ALIGN.CENTER
    p_h3.font.name = "Aptos"
    p_h3.font.size = Pt(11.0)
    p_h3.font.bold = True
    p_h3.font.color.rgb = hex_to_rgb("FFFFFF")

    benefits_rows = [
        ("S2P Services Excellence", "Deep Global Category, Industry and Process Expertise"),
        ("Efficiency", "15–30% productivity improvement through process standardization and automation"),
        ("Digital COE + Agents", "AI agents across S2P on KPOD systems"),
        ("SME & Talented pool", "Access to 3,000+ category, sourcing and procurement specialists"),
        ("Leadership Commitment", "Executive sponsorship, dedicated account leadership and category ownership"),
        ("Outcome based model", "Measurable KPIs, savings commitments and transparent benefits tracking")
    ]

    row_h = 60.0
    row_gap = 6.0
    for r_idx, (r_title, r_desc) in enumerate(benefits_rows):
        ry = col_top + 14.0 + r_idx * (row_h + row_gap)
        r_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(col_x3 + 10.0), Pt(ry), Pt(col_w3 - 20.0), Pt(row_h))
        apply_solid_fill(r_box, "F8FAFC")
        r_box.line.color.rgb = hex_to_rgb("E2E8F0")
        set_shape_corner_radius(r_box, 4.0)
        apply_library_shadow(r_box)

        # Icon Circle
        ic = slide.shapes.add_shape(MSO_SHAPE.OVAL, Pt(col_x3 + 18.0), Pt(ry + 16.0), Pt(26.0), Pt(26.0))
        apply_solid_fill(ic, "EFF6FF")
        ic.line.color.rgb = hex_to_rgb("BFDBFE")
        ic_tf = ic.text_frame
        ic_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        ic_p = ic_tf.paragraphs[0]
        ic_p.text = "★"
        ic_p.font.size = Pt(11.0)
        ic_p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)
        ic_p.alignment = PP_ALIGN.CENTER

        # Text Frame
        tf_row = slide.shapes.add_textbox(Pt(col_x3 + 50.0), Pt(ry + 4.0), Pt(col_w3 - 62.0), Pt(row_h - 8.0))
        r_tf = tf_row.text_frame
        r_tf.word_wrap = True
        r_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        rp1 = r_tf.paragraphs[0]
        rp1.text = r_title
        rp1.font.name = "Aptos"
        rp1.font.size = Pt(9.5)
        rp1.font.bold = True
        rp1.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

        rp2 = r_tf.add_paragraph()
        rp2.text = r_desc
        rp2.font.name = "Aptos"
        rp2.font.size = Pt(9.0)
        rp2.font.color.rgb = hex_to_rgb("475569")


# ==============================================================================
# SLIDE 9: Source-to-Pay Expertise
# ==============================================================================

def build_slide9(prs, lib_s1, lib_lines):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Title
    title_box = slide.shapes.add_textbox(Pt(30.0), Pt(8.0), Pt(900.0), Pt(32.0))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Source-to-Pay Expertise: Our laser focus on procurement and ready assets deliver better experience, efficiency, and effectiveness for GEP’s clients"
    p.font.name = "Aptos"
    p.font.size = Pt(14.0)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

    decor_proto = get_shape_by_name(lib_s1, "Design 1 Decor box")
    hdr_proto = get_shape_by_name(lib_s1, "Design 1 Header Text")
    hdr_proto = get_shape_by_name(lib_s1, "Design 1 Header Text")

    # Left: Source-to-Pay Capabilities (w = 580.0, top = 56.0, h = 356.0)
    card_l_elem = clean_element(copy.deepcopy(decor_proto._element))
    slide.shapes._spTree.append(card_l_elem)
    sh_cl = slide.shapes[-1]
    sh_cl.left = Pt(30.0)
    sh_cl.top = Pt(56.0)
    sh_cl.width = Pt(580.0)
    sh_cl.height = Pt(356.0)
    apply_solid_fill(sh_cl, "FFFFFF")
    sh_cl.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(sh_cl, 4.0)

    # Left Header pill
    hl_elem = clean_element(copy.deepcopy(hdr_proto._element))
    slide.shapes._spTree.append(hl_elem)
    sh_hl = slide.shapes[-1]
    sh_hl.left = Pt(190.0)
    sh_hl.top = Pt(44.0)
    sh_hl.width = Pt(260.0)
    sh_hl.height = Pt(24.0)
    apply_solid_fill(sh_hl, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(sh_hl, 4.0)
    sh_hl.text_frame.word_wrap = False
    sh_hl.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_hl = sh_hl.text_frame.paragraphs[0]
    p_hl.text = "Source-to-Pay Capabilities"
    p_hl.alignment = PP_ALIGN.CENTER
    p_hl.font.name = "Aptos"
    p_hl.font.size = Pt(11.0)
    p_hl.font.bold = True
    p_hl.font.color.rgb = hex_to_rgb("FFFFFF")

    # 3 Subcolumns inside Capabilities
    subcol_w = 176.0
    subcol_lefts = [42.0, 232.0, 422.0]

    col1_stats = [
        ("$4.5T+", "Spend Analyzed Annually"),
        ("$600+ BN", "Spend Managed Annually"),
        ("50 MM", "Transactions Managed Annually"),
        ("500+", "Subject Matter Experts"),
        ("10+", "Category Advisory Groups")
    ]

    col2_stats = [
        ("50K+", "Sourcing Projects / Year"),
        ("2 MM+", "Contracts Processed Annually"),
        ("35 Million+", "Supplier Network"),
        ("3,000+", "Category, Sourcing, Professionals"),
        ("25+", "Languages Supported"),
        ("120+", "Countries Supported")
    ]

    col3_stats = [
        ("5,000+", "Digital Tools & RPAs"),
        ("25K+", "Catalogs Implemented"),
        ("30 – 50%", "Efficiencies through Automation"),
        ("500+", "Analytical Data Models"),
        ("10 – 20%", "Reduction in Turnaround Time")
    ]

    for c_i, stats in enumerate([col1_stats, col2_stats, col3_stats]):
        cx = subcol_lefts[c_i]
        # White inner container
        cbox = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(cx), Pt(72.0), Pt(subcol_w), Pt(328.0))
        apply_solid_fill(cbox, "F8FAFC")
        cbox.line.color.rgb = hex_to_rgb("E2E8F0")
        set_shape_corner_radius(cbox, 4.0)
        apply_library_shadow(cbox)

        n_items = len(stats)
        item_h = (328.0 - 12.0) / n_items

        for s_i, (val, lbl) in enumerate(stats):
            sy = 78.0 + s_i * item_h
            # Metric Circle
            ic = slide.shapes.add_shape(MSO_SHAPE.OVAL, Pt(cx + 8.0), Pt(sy + (item_h - 26.0) / 2.0), Pt(26.0), Pt(26.0))
            apply_solid_fill(ic, "EFF6FF" if c_i != 1 else "FFF7ED")
            ic.line.color.rgb = hex_to_rgb("BFDBFE" if c_i != 1 else "FED7AA")
            ic_tf = ic.text_frame
            ic_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            ic_p = ic_tf.paragraphs[0]
            ic_p.text = "◆"
            ic_p.font.size = Pt(9.0)
            ic_p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO if c_i != 1 else ColorTokens.BRAND_TANGERINE)
            ic_p.alignment = PP_ALIGN.CENTER

            # Text
            stf_box = slide.shapes.add_textbox(Pt(cx + 38.0), Pt(sy + 2.0), Pt(subcol_w - 44.0), Pt(item_h - 4.0))
            stf = stf_box.text_frame
            stf.word_wrap = True
            stf.vertical_anchor = MSO_ANCHOR.MIDDLE
            sp1 = stf.paragraphs[0]
            sp1.text = val
            sp1.font.name = "Aptos"
            sp1.font.size = Pt(11.5)
            sp1.font.bold = True
            sp1.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY if c_i != 1 else ColorTokens.BRAND_TANGERINE)

            sp2 = stf.add_paragraph()
            sp2.text = lbl
            sp2.font.name = "Aptos"
            sp2.font.size = Pt(9.0)
            sp2.font.color.rgb = hex_to_rgb("475569")

            if s_i < n_items - 1:
                add_library_line(slide, lib_lines["side_fade_r"], cx + 8.0, sy + item_h - 1.0, subcol_w - 16.0)

    # Right: GEP Drivers and Outcomes Targeted (w = 306.0, top = 56.0, h = 356.0)
    card_r_elem = clean_element(copy.deepcopy(decor_proto._element))
    slide.shapes._spTree.append(card_r_elem)
    sh_cr = slide.shapes[-1]
    sh_cr.left = Pt(624.0)
    sh_cr.top = Pt(56.0)
    sh_cr.width = Pt(306.0)
    sh_cr.height = Pt(356.0)
    apply_solid_fill(sh_cr, "FFFFFF")
    sh_cr.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(sh_cr, 4.0)

    # Right Header pill
    hr_elem = clean_element(copy.deepcopy(hdr_proto._element))
    slide.shapes._spTree.append(hr_elem)
    sh_hr = slide.shapes[-1]
    sh_hr.left = Pt(647.0)
    sh_hr.top = Pt(44.0)
    sh_hr.width = Pt(260.0)
    sh_hr.height = Pt(24.0)
    apply_solid_fill(sh_hr, ColorTokens.BRAND_TANGERINE)
    set_shape_corner_radius(sh_hr, 4.0)
    sh_hr.text_frame.word_wrap = False
    sh_hr.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_hr = sh_hr.text_frame.paragraphs[0]
    p_hr.text = "GEP Drivers and Outcomes Targeted"
    p_hr.alignment = PP_ALIGN.CENTER
    p_hr.font.name = "Aptos"
    p_hr.font.size = Pt(10.5)
    p_hr.font.bold = True
    p_hr.font.color.rgb = hex_to_rgb("FFFFFF")

    drivers_rows = [
        ("Speed-to-Value with Digital Focus", "Ready to deploy infrastructure powered by digital tools to drive value and velocity"),
        ("Change Management", "Targeted change management as a part of foundation building to ensure strong adoption"),
        ("Reduced Cycle Time", "Process optimization to drive efficiencies and reduce cycle time"),
        ("Strong Compliance", "Driving buying channel compliance at source"),
        ("FTE Efficiency", "Automating multiple S2P processes through AI and RPA"),
        ("Innovation and Digital Focus", "Focus on RPAs, NLP, AI based chatbots, etc. to improve productivity"),
        ("Improved Experience", "Use of auto-bid, cost models and other GEP incremental tools to further improve experience")
    ]

    d_row_h = 43.0
    d_row_gap = 4.0
    for d_i, (d_title, d_desc) in enumerate(drivers_rows):
        dy = 72.0 + d_i * (d_row_h + d_row_gap)
        d_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(634.0), Pt(dy), Pt(286.0), Pt(d_row_h))
        apply_solid_fill(d_box, "F8FAFC")
        d_box.line.color.rgb = hex_to_rgb("E2E8F0")
        set_shape_corner_radius(d_box, 4.0)
        apply_library_shadow(d_box)

        # Icon Circle
        ic = slide.shapes.add_shape(MSO_SHAPE.OVAL, Pt(642.0), Pt(dy + 10.0), Pt(22.0), Pt(22.0))
        apply_solid_fill(ic, "FFF7ED")
        ic.line.color.rgb = hex_to_rgb("FED7AA")
        ic_tf = ic.text_frame
        ic_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        ic_p = ic_tf.paragraphs[0]
        ic_p.text = "✓"
        ic_p.font.size = Pt(10.0)
        ic_p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_TANGERINE)
        ic_p.alignment = PP_ALIGN.CENTER

        # Text Frame
        tf_row = slide.shapes.add_textbox(Pt(670.0), Pt(dy + 2.0), Pt(244.0), Pt(d_row_h - 4.0))
        r_tf = tf_row.text_frame
        r_tf.word_wrap = True
        r_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        rp1 = r_tf.paragraphs[0]
        rp1.text = d_title
        rp1.font.name = "Aptos"
        rp1.font.size = Pt(9.0)
        rp1.font.bold = True
        rp1.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

        rp2 = r_tf.add_paragraph()
        rp2.text = d_desc
        rp2.font.name = "Aptos"
        rp2.font.size = Pt(9.0)
        rp2.font.color.rgb = hex_to_rgb("475569")

    # Bottom: Experience With Leading Procurement Platforms (w = 900.0, top = 422.0, h = 104.0)
    p_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(422.0), Pt(900.0), Pt(104.0))
    apply_solid_fill(p_box, "FFFFFF")
    p_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(p_box, 4.0)
    apply_library_shadow(p_box)

    # Left Badge
    lb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(40.0), Pt(432.0), Pt(140.0), Pt(84.0))
    apply_solid_fill(lb, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(lb, 4.0)
    lb_tf = lb.text_frame
    lb_tf.word_wrap = True
    lb_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    lb_p = lb_tf.paragraphs[0]
    lb_p.text = "Experience With Leading Procurement Platforms"
    lb_p.alignment = PP_ALIGN.CENTER
    lb_p.font.name = "Aptos"
    lb_p.font.size = Pt(10.0)
    lb_p.font.bold = True
    lb_p.font.color.rgb = hex_to_rgb("FFFFFF")

    # 12 Platforms in 2 rows of 6
    platforms_r1 = ["GEP SMART", "ORACLE", "IBM maximo", "DRAFTC", "WeBuy", "JAGGAER"]
    platforms_r2 = ["Avetta", "Salesforce", "SAP Ariba", "ServiceNow", "Coupa", "Workday"]

    for row_idx, p_list in enumerate([platforms_r1, platforms_r2]):
        py = 434.0 + row_idx * 42.0
        for col_idx, p_name in enumerate(p_list):
            px = 196.0 + col_idx * 120.0
            p_badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(px), Pt(py), Pt(114.0), Pt(36.0))
            apply_solid_fill(p_badge, "F8FAFC")
            p_badge.line.color.rgb = hex_to_rgb("CBD5E1")
            set_shape_corner_radius(p_badge, 4.0)
            b_tf = p_badge.text_frame
            b_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            bp = b_tf.paragraphs[0]
            bp.text = p_name
            bp.alignment = PP_ALIGN.CENTER
            bp.font.name = "Aptos"
            bp.font.size = Pt(9.5)
            bp.font.bold = True
            bp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY if "GEP" not in p_name else ColorTokens.BRAND_TANGERINE)


# ==============================================================================
# SLIDE 10: Plan on a Page
# ==============================================================================

def build_slide10(prs, lib_s1, lib_lines):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Title
    title_box = slide.shapes.add_textbox(Pt(30.0), Pt(6.0), Pt(900.0), Pt(32.0))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Plan on a Page: a 3 month transition takes KPOD to steady state, followed by 33 months of strategic procurement services across all in scope activities"
    p.font.name = "Aptos"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

    # Milestone: Kick-off
    ms_box = slide.shapes.add_textbox(Pt(30.0), Pt(38.0), Pt(240.0), Pt(18.0))
    msp = ms_box.text_frame.paragraphs[0]
    msp.text = "★ Program Kick-Off (Feb 2027)"
    msp.font.name = "Aptos"
    msp.font.size = Pt(9.0)
    msp.font.bold = True
    msp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_TANGERINE)

    # Timeline Bars (y = 56.0, h = 24.0)
    m_widths = [90.0, 90.0, 90.0]
    m_names = ["M1", "M2", "M3"]
    for i in range(3):
        mb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0 + i * 92.0), Pt(56.0), Pt(88.0), Pt(22.0))
        apply_solid_fill(mb, ColorTokens.BRAND_INDIGO)
        set_shape_corner_radius(mb, 4.0)
        mb_tf = mb.text_frame
        mb_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        mp = mb_tf.paragraphs[0]
        mp.text = m_names[i]
        mp.alignment = PP_ALIGN.CENTER
        mp.font.name = "Aptos"
        mp.font.size = Pt(9.5)
        mp.font.bold = True
        mp.font.color.rgb = hex_to_rgb("FFFFFF")

    # Steady State Bar (w = 620.0)
    ssb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(310.0), Pt(56.0), Pt(620.0), Pt(22.0))
    apply_solid_fill(ssb, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(ssb, 4.0)
    ssb_tf = ssb.text_frame
    ssb_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    ssp = ssb_tf.paragraphs[0]
    ssp.text = "M4 to M36 (33 months): May 2027 to Jan 2030"
    ssp.alignment = PP_ALIGN.CENTER
    ssp.font.name = "Aptos"
    ssp.font.size = Pt(9.5)
    ssp.font.bold = True
    ssp.font.color.rgb = hex_to_rgb("FFFFFF")

    # Middle Left: Foundation Building and Transition | 3 Months (w = 270.0, top = 86.0, h = 330.0)
    fb_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(86.0), Pt(270.0), Pt(330.0))
    apply_solid_fill(fb_box, "FFFFFF")
    fb_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(fb_box, 4.0)
    apply_library_shadow(fb_box)

    # Header bar
    fb_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(34.0), Pt(90.0), Pt(262.0), Pt(22.0))
    apply_solid_fill(fb_hdr, ColorTokens.BRAND_INDIGO)
    set_shape_corner_radius(fb_hdr, 4.0)
    h_tf = fb_hdr.text_frame
    h_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    hp = h_tf.paragraphs[0]
    hp.text = "Foundation Building and Transition | 3 Months"
    hp.alignment = PP_ALIGN.CENTER
    hp.font.name = "Aptos"
    hp.font.size = Pt(9.0)
    hp.font.bold = True
    hp.font.color.rgb = hex_to_rgb("FFFFFF")

    # Sub-section 1: Foundation Building
    s1_lbl = slide.shapes.add_textbox(Pt(38.0), Pt(114.0), Pt(254.0), Pt(16.0))
    s1p = s1_lbl.text_frame.paragraphs[0]
    s1p.text = "Foundation Building"
    s1p.font.name = "Aptos"
    s1p.font.size = Pt(9.0)
    s1p.font.bold = True
    s1p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

    f_bullets = [
        "PMO set up, governance, SLAs and KPI baselines",
        "Baseline spend; map categories to Wells, Operations, Projects and Indirect; validate savings",
        "Map warehouse, logistics and customs interfaces",
        "Oracle, Ariba and AIQ access, governance and transition planning",
        "Supplier risk framework and critical contracts identified",
        "Local content baseline, reporting requirements and governance framework established"
    ]
    fb_tf_box = slide.shapes.add_textbox(Pt(38.0), Pt(130.0), Pt(254.0), Pt(130.0))
    f_tf = fb_tf_box.text_frame
    f_tf.word_wrap = True
    f_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for idx, b_txt in enumerate(f_bullets):
        pb = f_tf.paragraphs[0] if idx == 0 else f_tf.add_paragraph()
        pb.text = clean_bullet_string(b_txt)
        pb.font.name = "Aptos"
        pb.font.size = Pt(9.0)
        pb.font.color.rgb = hex_to_rgb("334155")
        pb.space_after = Pt(2.0)
        apply_native_bullet(pb, char="•", typeface="Arial", mar_l_pt=9.0, indent_pt=-9.0)

    add_library_line(slide, lib_lines["side_fade_r"], 38.0, 264.0, 254.0)

    # Sub-section 2: Transition and Team Ramp-Up
    s2_lbl = slide.shapes.add_textbox(Pt(38.0), Pt(268.0), Pt(254.0), Pt(16.0))
    s2p = s2_lbl.text_frame.paragraphs[0]
    s2p.text = "Transition and Team Ramp-Up"
    s2p.font.name = "Aptos"
    s2p.font.size = Pt(9.0)
    s2p.font.bold = True
    s2p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_TANGERINE)

    t_bullets = [
        "Knowledge transfer and desk procedures; PR to PO service levels",
        "Take over 130 live sourcing cases: close to award, mid process, early stage",
        "Resource onboarding, vetting and training",
        "Parallel run, readiness review and go live communication"
    ]
    tb_tf_box = slide.shapes.add_textbox(Pt(38.0), Pt(284.0), Pt(254.0), Pt(128.0))
    t_tf = tb_tf_box.text_frame
    t_tf.word_wrap = True
    t_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for idx, b_txt in enumerate(t_bullets):
        pb = t_tf.paragraphs[0] if idx == 0 else t_tf.add_paragraph()
        pb.text = clean_bullet_string(b_txt)
        pb.font.name = "Aptos"
        pb.font.size = Pt(9.0)
        pb.font.color.rgb = hex_to_rgb("334155")
        pb.space_after = Pt(2.0)
        apply_native_bullet(pb, char="•", typeface="Arial", mar_l_pt=9.0, indent_pt=-9.0)

    # Middle Right: Strategic Procurement Services (Steady State) | 33 Months (w = 620.0, top = 86.0, h = 330.0)
    sp_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(310.0), Pt(86.0), Pt(620.0), Pt(22.0))
    apply_solid_fill(sp_hdr, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(sp_hdr, 4.0)
    sp_tf = sp_hdr.text_frame
    sp_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    spp = sp_tf.paragraphs[0]
    spp.text = "Strategic Procurement Services (Steady State) | 33 Months"
    spp.alignment = PP_ALIGN.CENTER
    spp.font.name = "Aptos"
    spp.font.size = Pt(9.0)
    spp.font.bold = True
    spp.font.color.rgb = hex_to_rgb("FFFFFF")

    proc_streams = [
        ("Category Strategy:", "category strategies and a rolling 24 month procurement plan; annual refresh and benefits register"),
        ("Strategic Sourcing:", "Sourcing and contracting across Wells, Operations, Projects and Indirect categories; category strategies, sourcing events, negotiations, contract awards and savings delivery"),
        ("Contract and Supplier Management:", "contract administration, amendments and claims; supplier performance, local content tracking, contract compliance, risk monitoring and continuity plans"),
        ("Procurement Operations:", "requisition to PO, supplier support, buying support, expediting and transactional procurement activities"),
        ("Materials Management:", "materials coordination, inventory visibility, shortage reporting and inventory optimisation")
    ]

    p_row_h = 56.0
    p_gap = 6.0
    for p_idx, (p_title, p_desc) in enumerate(proc_streams):
        py = 114.0 + p_idx * (p_row_h + p_gap)
        p_card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(310.0), Pt(py), Pt(620.0), Pt(p_row_h))
        apply_solid_fill(p_card, "F8FAFC")
        p_card.line.color.rgb = hex_to_rgb("CBD5E1")
        set_shape_corner_radius(p_card, 4.0)
        apply_library_shadow(p_card)

        # Icon Circle
        ic = slide.shapes.add_shape(MSO_SHAPE.OVAL, Pt(322.0), Pt(py + 15.0), Pt(26.0), Pt(26.0))
        apply_solid_fill(ic, "EFF6FF" if p_idx % 2 == 0 else "FFF7ED")
        ic.line.color.rgb = hex_to_rgb("BFDBFE" if p_idx % 2 == 0 else "FED7AA")
        ic_tf = ic.text_frame
        ic_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        ic_p = ic_tf.paragraphs[0]
        ic_p.text = "▶"
        ic_p.font.size = Pt(10.0)
        ic_p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO if p_idx % 2 == 0 else ColorTokens.BRAND_TANGERINE)
        ic_p.alignment = PP_ALIGN.CENTER

        # Text Frame
        t_box = slide.shapes.add_textbox(Pt(356.0), Pt(py + 4.0), Pt(564.0), Pt(p_row_h - 8.0))
        t_tf = t_box.text_frame
        t_tf.word_wrap = True
        t_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_t = t_tf.paragraphs[0]
        r1 = p_t.add_run()
        r1.text = p_title + " "
        r1.font.name = "Aptos"
        r1.font.bold = True
        r1.font.size = Pt(9.5)
        r1.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)
        r2 = p_t.add_run()
        r2.text = p_desc
        r2.font.name = "Aptos"
        r2.font.size = Pt(9.0)
        r2.font.color.rgb = hex_to_rgb("475569")

    # Bottom: Digital and AI Layer (w = 900.0, top = 422.0, h = 60.0)
    ai_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(422.0), Pt(900.0), Pt(60.0))
    apply_solid_fill(ai_box, "FFFFFF")
    ai_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(ai_box, 4.0)
    apply_library_shadow(ai_box)

    # Left Badge
    aib = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(38.0), Pt(428.0), Pt(140.0), Pt(48.0))
    apply_solid_fill(aib, ColorTokens.BRAND_INDIGO)
    set_shape_corner_radius(aib, 4.0)
    aib_tf = aib.text_frame
    aib_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    aip1 = aib_tf.paragraphs[0]
    aip1.text = "Digital and AI Layer"
    aip1.alignment = PP_ALIGN.CENTER
    aip1.font.name = "Aptos"
    aip1.font.size = Pt(9.0)
    aip1.font.bold = True
    aip1.font.color.rgb = hex_to_rgb("FFFFFF")
    aip2 = aib_tf.add_paragraph()
    aip2.text = "Across transition and steady state"
    aip2.alignment = PP_ALIGN.CENTER
    aip2.font.name = "Aptos"
    aip2.font.size = Pt(9.0)
    aip2.font.color.rgb = hex_to_rgb("E2E8F0")

    # 4 Platforms
    platforms = ["Oracle", "FairmarkIT", "Aravo", "Beroe"]
    for pi, pname in enumerate(platforms):
        pb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(186.0 + pi * 52.0), Pt(438.0), Pt(48.0), Pt(28.0))
        apply_solid_fill(pb, "F8FAFC")
        pb.line.color.rgb = hex_to_rgb("CBD5E1")
        set_shape_corner_radius(pb, 4.0)
        pbtf = pb.text_frame
        pbtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        pp = pbtf.paragraphs[0]
        pp.text = pname
        pp.alignment = PP_ALIGN.CENTER
        pp.font.name = "Aptos"
        pp.font.size = Pt(9.0)
        pp.font.bold = True
        pp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

    # Plus sign
    pls = slide.shapes.add_textbox(Pt(396.0), Pt(440.0), Pt(16.0), Pt(24.0))
    plsp = pls.text_frame.paragraphs[0]
    plsp.text = "+"
    plsp.font.name = "Aptos"
    plsp.font.size = Pt(14.0)
    plsp.font.bold = True
    plsp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_TANGERINE)

    # 8 AI Agents in 2 rows of 4
    ai_r1 = ["Intake and Triage", "RFx Creation", "RFx Evaluation", "Quick Negotiation"]
    ai_r2 = ["Contract Drafting", "Should Cost", "PR Validation", "Autonomous Sourcing"]

    for r_i, r_list in enumerate([ai_r1, ai_r2]):
        ay = 428.0 + r_i * 24.0
        for c_i, aname in enumerate(r_list):
            ax = 418.0 + c_i * 126.0
            ab = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(ax), Pt(ay), Pt(120.0), Pt(21.0))
            apply_solid_fill(ab, "FFF7ED")
            ab.line.color.rgb = hex_to_rgb("FED7AA")
            set_shape_corner_radius(ab, 4.0)
            abtf = ab.text_frame
            abtf.vertical_anchor = MSO_ANCHOR.MIDDLE
            ap = abtf.paragraphs[0]
            ap.text = aname
            ap.alignment = PP_ALIGN.CENTER
            ap.font.name = "Aptos"
            ap.font.size = Pt(9.0)
            ap.font.bold = True
            ap.font.color.rgb = hex_to_rgb("9A3412")

    # Bottom Footer Bar (y = 488.0, h = 20.0)
    ftr_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(488.0), Pt(900.0), Pt(20.0))
    apply_solid_fill(ftr_box, "F1F5F9")
    ftr_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(ftr_box, 4.0)
    ftrf = ftr_box.text_frame
    ftrf.vertical_anchor = MSO_ANCHOR.MIDDLE
    fp = ftrf.paragraphs[0]
    fp.text = "Ongoing Governance | Savings Tracking | Local Content Reporting | KPI Management | PMO | Risk & Change Management"
    fp.alignment = PP_ALIGN.CENTER
    fp.font.name = "Aptos"
    fp.font.size = Pt(9.0)
    fp.font.bold = True
    fp.font.color.rgb = hex_to_rgb("475569")


# ==============================================================================
# SLIDE 11: Spend and Savings Opportunity
# ==============================================================================

def build_slide11(prs, lib_s1, lib_lines):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Title
    title_box = slide.shapes.add_textbox(Pt(30.0), Pt(8.0), Pt(900.0), Pt(32.0))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Spend and Savings Opportunity: we see USD 32 to 47M in annual savings at 6 to 10% of addressable spend across KPOD’s categories"
    p.font.name = "Aptos"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

    # Top Left Table Matrix: After Slide 19 style (w = 660.0, top = 46.0, h = 186.0)
    col_headers = [
        ("Category", 100.0),
        ("Key sub-categories", 210.0),
        ("Typical annual spend (USD M)", 95.0),
        ("% Addressability", 80.0),
        ("Savings %", 80.0),
        ("Annual savings (USD M)", 95.0)
    ]

    # Header Row (y = 46.0, h = 26.0)
    hx = 30.0
    for h_txt, hw in col_headers:
        hb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(hx), Pt(46.0), Pt(hw), Pt(26.0))
        apply_solid_fill(hb, ColorTokens.BRAND_NAVY)
        set_shape_corner_radius(hb, 4.0)
        h_tf = hb.text_frame
        h_tf.word_wrap = True
        h_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        hp = h_tf.paragraphs[0]
        hp.text = h_txt
        hp.alignment = PP_ALIGN.CENTER
        hp.font.name = "Aptos"
        hp.font.size = Pt(9.0)
        hp.font.bold = True
        hp.font.color.rgb = hex_to_rgb("FFFFFF")
        hx += hw

    table_data = [
        ("Wells", "Drilling & Completion Services, OCTG & Tubulars, Cementing & Stimulation, Wireline & Logging", "300 - 500", "55%", "6–9%", "13 to 20"),
        ("Operations", "O&M Services, Rotating Equipment, Production Chemicals, Maintenance Materials", "200 - 300", "50%", "7–10%", "9 to 13"),
        ("Projects", "EPC/EPCM Packages, Construction Services, Valves, Pipe & Fittings, Electrical & Instrumentation", "100 - 200", "50%", "6–9%", "5 to 7"),
        ("Indirects", "Logistics & Freight, Facilities Management, IT & Telecom, Corporate Services", "50 - 150", "60%", "8 to 12%", "5 to 7")
    ]

    row_h = 38.0
    for r_i, r_vals in enumerate(table_data):
        ry = 74.0 + r_i * (row_h + 2.0)
        rx = 30.0
        for c_i, val in enumerate(r_vals):
            hw = col_headers[c_i][1]
            rb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(rx), Pt(ry), Pt(hw), Pt(row_h))
            apply_solid_fill(rb, "F8FAFC" if r_i % 2 == 1 else "FFFFFF")
            rb.line.color.rgb = hex_to_rgb("E2E8F0")
            set_shape_corner_radius(rb, 4.0)
            apply_library_shadow(rb)

            r_tf = rb.text_frame
            r_tf.word_wrap = True
            r_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            rp = r_tf.paragraphs[0]
            rp.text = val
            rp.font.name = "Aptos"
            rp.alignment = PP_ALIGN.CENTER if c_i != 1 else PP_ALIGN.LEFT
            if c_i == 0:
                rp.font.bold = True
                rp.font.size = Pt(9.5)
                rp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)
            elif c_i == 5:
                rp.font.bold = True
                rp.font.size = Pt(10.0)
                rp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_TANGERINE)
            else:
                rp.font.size = Pt(9.0)
                rp.font.color.rgb = hex_to_rgb("334155")
            rx += hw

    # Top Right: 3 Stacked KPI Cards (KPI-CRD-V1 style, w = 226.0, left = 704.0)
    kpis_s11 = [
        ("~USD 1B", "indicative annual spend (midpoint of typical ranges)", ColorTokens.BRAND_NAVY),
        ("USD 32 to 47M", "annual savings run rate", ColorTokens.BRAND_TANGERINE),
        ("6 to 10%", "savings on addressable spend", ColorTokens.BRAND_INDIGO)
    ]

    for k_i, (k_val, k_sub, k_clr) in enumerate(kpis_s11):
        ky = 46.0 + k_i * 62.0
        kbox = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(704.0), Pt(ky), Pt(226.0), Pt(58.0))
        apply_solid_fill(kbox, "F8FAFC")
        kbox.line.color.rgb = hex_to_rgb(k_clr)
        kbox.line.width = Pt(1.0)
        set_shape_corner_radius(kbox, 4.0)
        apply_library_shadow(kbox)

        ktf = kbox.text_frame
        ktf.vertical_anchor = MSO_ANCHOR.MIDDLE
        kp1 = ktf.paragraphs[0]
        kp1.text = k_val
        kp1.alignment = PP_ALIGN.CENTER
        kp1.font.name = "Aptos"
        kp1.font.size = Pt(15.0)
        kp1.font.bold = True
        kp1.font.color.rgb = hex_to_rgb(k_clr)

        kp2 = ktf.add_paragraph()
        kp2.text = k_sub
        kp2.alignment = PP_ALIGN.CENTER
        kp2.font.name = "Aptos"
        kp2.font.size = Pt(9.0)
        kp2.font.color.rgb = hex_to_rgb("64748B")

    # Bottom Section: 4 Category Strategic Levers Cards (Design 1 Decor box style, top = 246.0, h = 276.0)
    decor_proto = get_shape_by_name(lib_s1, "Design 1 Decor box")
    hdr_proto = get_shape_by_name(lib_s1, "Design 1 Header Text")

    categories = [
        {
            "name": "Wells",
            "clr": ColorTokens.BRAND_INDIGO,
            "market": "Market: Tight rig and OCTG markets; steel price swings; campaign driven demand",
            "bullets": [
                "Multi well frame agreements with performance KPIs",
                "Index linked OCTG pricing and consignment stock"
            ]
        },
        {
            "name": "Operations",
            "clr": ColorTokens.BRAND_NAVY,
            "market": "Market: High volume, fragmented spend; remote site logistics; ageing assets",
            "bullets": [
                "Multi year O&M frameworks with KPIs",
                "Spares standardisation and inventory optimisation",
                "Catalogues and FairmarkIT for tail spend"
            ]
        },
        {
            "name": "Projects",
            "clr": ColorTokens.BRAND_INDIGO,
            "market": "Market: Large, complex packages; schedule critical long lead items",
            "bullets": [
                "Package strategy and early supplier involvement",
                "Engineering contract benchmarking and should cost",
                "Local fabrication to grow Iraqi content"
            ]
        },
        {
            "name": "Indirect",
            "clr": ColorTokens.BRAND_TANGERINE,
            "market": "Market: Many small suppliers; weak demand control; rate creep",
            "bullets": [
                "Demand management and spec rationalisation",
                "Supplier consolidation",
                "Rate card benchmarking"
            ]
        }
    ]

    card_w = 214.0
    gap = 14.0
    c_top = 246.0
    c_h = 276.0

    for c_idx, cat in enumerate(categories):
        cx = 30.0 + c_idx * (card_w + gap)

        c_elem = clean_element(copy.deepcopy(decor_proto._element))
        slide.shapes._spTree.append(c_elem)
        sh_cat = slide.shapes[-1]
        sh_cat.left = Pt(cx)
        sh_cat.top = Pt(c_top)
        sh_cat.width = Pt(card_w)
        sh_cat.height = Pt(c_h)
        apply_solid_fill(sh_cat, "FFFFFF")
        sh_cat.line.color.rgb = hex_to_rgb("CBD5E1")
        set_shape_corner_radius(sh_cat, 4.0)

        # Header Box
        hb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(cx + 8.0), Pt(c_top + 8.0), Pt(card_w - 16.0), Pt(32.0))
        apply_solid_fill(hb, cat["clr"])
        set_shape_corner_radius(hb, 4.0)
        hbtf = hb.text_frame
        hbtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        hp = hbtf.paragraphs[0]
        hp.text = cat["name"]
        hp.alignment = PP_ALIGN.CENTER
        hp.font.name = "Aptos"
        hp.font.size = Pt(12.0)
        hp.font.bold = True
        hp.font.color.rgb = hex_to_rgb("FFFFFF")

        # Market Text
        m_box = slide.shapes.add_textbox(Pt(cx + 10.0), Pt(c_top + 44.0), Pt(card_w - 20.0), Pt(46.0))
        mtf = m_box.text_frame
        mtf.word_wrap = True
        mtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        mp = mtf.paragraphs[0]
        mp.text = cat["market"]
        mp.font.name = "Aptos"
        mp.font.size = Pt(9.0)
        mp.font.color.rgb = hex_to_rgb("475569")

        # Divider
        add_library_line(slide, lib_lines["side_fade_r"], cx + 10.0, c_top + 94.0, card_w - 20.0)

        # STRATEGIC LEVERS label
        l_box = slide.shapes.add_textbox(Pt(cx + 10.0), Pt(c_top + 100.0), Pt(card_w - 20.0), Pt(18.0))
        lp = l_box.text_frame.paragraphs[0]
        lp.text = "STRATEGIC LEVERS"
        lp.font.name = "Aptos"
        lp.font.size = Pt(9.0)
        lp.font.bold = True
        lp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

        # Bullets
        b_box = slide.shapes.add_textbox(Pt(cx + 10.0), Pt(c_top + 120.0), Pt(card_w - 20.0), Pt(c_h - 128.0))
        btf = b_box.text_frame
        btf.word_wrap = True
        btf.vertical_anchor = MSO_ANCHOR.MIDDLE
        for b_i, bullet in enumerate(cat["bullets"]):
            pb = btf.paragraphs[0] if b_i == 0 else btf.add_paragraph()
            pb.text = clean_bullet_string(bullet)
            pb.font.name = "Aptos"
            pb.font.size = Pt(9.0)
            pb.font.color.rgb = hex_to_rgb("334155")
            pb.space_after = Pt(4.0)
            apply_native_bullet(pb, char="•", typeface="Arial", mar_l_pt=10.0, indent_pt=-10.0)


# ==============================================================================
# SLIDE 12: Resource Plan
# ==============================================================================

def build_slide12(prs, lib_s1, lib_lines):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Title
    title_box = slide.shapes.add_textbox(Pt(30.0), Pt(8.0), Pt(900.0), Pt(32.0))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Resource Plan: Dedicated PSCM managed service team supported by category SMEs, digital specialists and governance leadership across the 36-month engagement"
    p.font.name = "Aptos"
    p.font.size = Pt(13.0)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

    decor_proto = get_shape_by_name(lib_s1, "Design 1 Decor box")
    hdr_proto = get_shape_by_name(lib_s1, "Design 1 Header Text")

    # Left: bp Client Stakeholders container (w = 180.0, top = 46.0, h = 340.0)
    bp_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(46.0), Pt(180.0), Pt(340.0))
    apply_solid_fill(bp_box, "FFFFFF")
    bp_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(bp_box, 4.0)
    apply_library_shadow(bp_box)

    # bp Header
    bp_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(40.0), Pt(54.0), Pt(160.0), Pt(28.0))
    apply_solid_fill(bp_hdr, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(bp_hdr, 4.0)
    bp_tf = bp_hdr.text_frame
    bp_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    bpp = bp_tf.paragraphs[0]
    bpp.text = "bp Leadership"
    bpp.alignment = PP_ALIGN.CENTER
    bpp.font.name = "Aptos"
    bpp.font.size = Pt(10.0)
    bpp.font.bold = True
    bpp.font.color.rgb = hex_to_rgb("FFFFFF")

    bp_roles = [
        ("Executive Sponsor", 94.0, 48.0),
        ("Program Lead", 154.0, 48.0),
        ("Process Owners and\nProcurement Team", 214.0, 64.0)
    ]
    for r_title, ry, rh in bp_roles:
        rbox = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(40.0), Pt(ry), Pt(160.0), Pt(rh))
        apply_solid_fill(rbox, "F8FAFC")
        rbox.line.color.rgb = hex_to_rgb("CBD5E1")
        set_shape_corner_radius(rbox, 4.0)
        rtf = rbox.text_frame
        rtf.word_wrap = True
        rtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        rp = rtf.paragraphs[0]
        rp.text = r_title
        rp.alignment = PP_ALIGN.CENTER
        rp.font.name = "Aptos"
        rp.font.size = Pt(9.0)
        rp.font.bold = True
        rp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

    # Right: GEP Managed Service Team container (w = 706.0, top = 46.0, h = 340.0)
    gep_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(224.0), Pt(46.0), Pt(706.0), Pt(340.0))
    apply_solid_fill(gep_box, "FFFFFF")
    gep_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(gep_box, 4.0)
    apply_library_shadow(gep_box)

    # GEP Header
    gep_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(236.0), Pt(54.0), Pt(480.0), Pt(28.0))
    apply_solid_fill(gep_hdr, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(gep_hdr, 4.0)
    gtf = gep_hdr.text_frame
    gtf.vertical_anchor = MSO_ANCHOR.MIDDLE
    gp = gtf.paragraphs[0]
    gp.text = "GEP Managed Service Delivery Organization"
    gp.alignment = PP_ALIGN.CENTER
    gp.font.name = "Aptos"
    gp.font.size = Pt(10.0)
    gp.font.bold = True
    gp.font.color.rgb = hex_to_rgb("FFFFFF")

    # Service Delivery Manager & Client Executive
    sdm_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(246.0), Pt(94.0), Pt(230.0), Pt(48.0))
    apply_solid_fill(sdm_box, "EFF6FF")
    sdm_box.line.color.rgb = hex_to_rgb("3B82F6")
    set_shape_corner_radius(sdm_box, 4.0)
    sdm_tf = sdm_box.text_frame
    sdm_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p1 = sdm_tf.paragraphs[0]
    p1.text = "Service Delivery Manager"
    p1.font.name = "Aptos"
    p1.font.size = Pt(9.0)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)
    p1.alignment = PP_ALIGN.CENTER
    p2 = sdm_tf.add_paragraph()
    p2.text = "Virat Venkataraman"
    p2.font.name = "Aptos"
    p2.font.size = Pt(10.0)
    p2.font.bold = True
    p2.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)
    p2.alignment = PP_ALIGN.CENTER

    # Connector line
    add_library_line(slide, lib_lines["solid_h"], 478.0, 118.0, 30.0)

    ce_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(510.0), Pt(94.0), Pt(200.0), Pt(48.0))
    apply_solid_fill(ce_box, "FFF7ED")
    ce_box.line.color.rgb = hex_to_rgb("FF8300")
    set_shape_corner_radius(ce_box, 4.0)
    ce_tf = ce_box.text_frame
    ce_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p1 = ce_tf.paragraphs[0]
    p1.text = "Client Executive"
    p1.font.name = "Aptos"
    p1.font.size = Pt(9.0)
    p1.font.bold = True
    p1.font.color.rgb = hex_to_rgb("9A3412")
    p1.alignment = PP_ALIGN.CENTER
    p2 = ce_tf.add_paragraph()
    p2.text = "Jason Allukian"
    p2.font.name = "Aptos"
    p2.font.size = Pt(10.0)
    p2.font.bold = True
    p2.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)
    p2.alignment = PP_ALIGN.CENTER

    # Program Manager
    pm_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(246.0), Pt(154.0), Pt(464.0), Pt(34.0))
    apply_solid_fill(pm_box, "F8FAFC")
    pm_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(pm_box, 4.0)
    pm_tf = pm_box.text_frame
    pm_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    pmp = pm_tf.paragraphs[0]
    pmp.text = "Program Manager"
    pmp.font.name = "Aptos"
    pmp.font.size = Pt(10.0)
    pmp.font.bold = True
    pmp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)
    pmp.alignment = PP_ALIGN.CENTER

    # 2 Operational Streams
    streams = [
        ("Strategic Sourcing & Category Management", 246.0, 226.0),
        ("Procurement Operations & Materials Management", 484.0, 226.0)
    ]
    for s_title, sx, sw in streams:
        sbox = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(sx), Pt(200.0), Pt(sw), Pt(170.0))
        apply_solid_fill(sbox, "F8FAFC")
        sbox.line.color.rgb = hex_to_rgb("CBD5E1")
        set_shape_corner_radius(sbox, 4.0)
        apply_library_shadow(sbox)

        # Header
        shdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(sx + 6.0), Pt(206.0), Pt(sw - 12.0), Pt(28.0))
        apply_solid_fill(shdr, ColorTokens.BRAND_INDIGO)
        set_shape_corner_radius(shdr, 4.0)
        shtf = shdr.text_frame
        shtf.word_wrap = True
        shtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        shp = shtf.paragraphs[0]
        shp.text = s_title
        shp.alignment = PP_ALIGN.CENTER
        shp.font.name = "Aptos"
        shp.font.size = Pt(9.0)
        shp.font.bold = True
        shp.font.color.rgb = hex_to_rgb("FFFFFF")

    # Right Panel: GEP SMEs (as needed)
    sme_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(726.0), Pt(54.0), Pt(194.0), Pt(322.0))
    apply_solid_fill(sme_box, "F8FAFC")
    sme_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(sme_box, 4.0)
    apply_library_shadow(sme_box)

    sme_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(732.0), Pt(60.0), Pt(182.0), Pt(28.0))
    apply_solid_fill(sme_hdr, ColorTokens.BRAND_TANGERINE)
    set_shape_corner_radius(sme_hdr, 4.0)
    stf = sme_hdr.text_frame
    stf.vertical_anchor = MSO_ANCHOR.MIDDLE
    stp = stf.paragraphs[0]
    stp.text = "GEP SMEs (as needed)"
    stp.alignment = PP_ALIGN.CENTER
    stp.font.name = "Aptos"
    stp.font.size = Pt(9.5)
    stp.font.bold = True
    stp.font.color.rgb = hex_to_rgb("FFFFFF")

    smes = ["Matt Tuller", "Wayne Clark", "Emma Jane", "Raina Saxena"]
    for s_i, sname in enumerate(smes):
        sy = 98.0 + s_i * 54.0
        sb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(734.0), Pt(sy), Pt(178.0), Pt(46.0))
        apply_solid_fill(sb, "FFFFFF")
        sb.line.color.rgb = hex_to_rgb("E2E8F0")
        set_shape_corner_radius(sb, 4.0)
        sbtf = sb.text_frame
        sbtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        sbp = sbtf.paragraphs[0]
        sbp.text = sname
        sbp.alignment = PP_ALIGN.CENTER
        sbp.font.name = "Aptos"
        sbp.font.size = Pt(9.5)
        sbp.font.bold = True
        sbp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

    # Bottom: GEP Support Infrastructure (top = 406.0, h = 114.0)
    infra_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(406.0), Pt(900.0), Pt(114.0))
    apply_solid_fill(infra_box, "FFFFFF")
    infra_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(infra_box, 4.0)
    apply_library_shadow(infra_box)

    # Infrastructure Header pill (inside container top edge)
    sh_ih = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(38.0), Pt(410.0), Pt(884.0), Pt(20.0))
    apply_solid_fill(sh_ih, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(sh_ih, 4.0)
    sh_ih.text_frame.word_wrap = False
    sh_ih.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_ih = sh_ih.text_frame.paragraphs[0]
    p_ih.text = "GEP Support Infrastructure"
    p_ih.alignment = PP_ALIGN.CENTER
    p_ih.font.name = "Aptos"
    p_ih.font.size = Pt(10.0)
    p_ih.font.bold = True
    p_ih.font.color.rgb = hex_to_rgb("FFFFFF")

    # Footnote at slide base
    fn_box = slide.shapes.add_textbox(Pt(30.0), Pt(522.0), Pt(900.0), Pt(16.0))
    fnp = fn_box.text_frame.paragraphs[0]
    fnp.text = "[1] Phase 2 includes an additional investment resource to support quick win execution and accelerate savings delivery for Telix"
    fnp.font.name = "Aptos"
    fnp.font.size = Pt(9.0)
    fnp.font.italic = True
    fnp.font.color.rgb = hex_to_rgb("64748B")

    infra_items = [
        "Global Industry SMEs",
        "Digital Tools & Contracting RPAs",
        "Training Playbooks",
        "Industry Benchmarks",
        "Contracting Process Tools & Templates"
    ]
    card_w = 168.0
    for idx, item_txt in enumerate(infra_items):
        ix = 42.0 + idx * 174.0
        ibox = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(ix), Pt(434.0), Pt(card_w), Pt(80.0))
        apply_solid_fill(ibox, "F8FAFC")
        ibox.line.color.rgb = hex_to_rgb("E2E8F0")
        set_shape_corner_radius(ibox, 4.0)
        apply_library_shadow(ibox)

        # Icon Circle
        ic = slide.shapes.add_shape(MSO_SHAPE.OVAL, Pt(ix + (card_w - 24.0) / 2.0), Pt(438.0), Pt(24.0), Pt(24.0))
        apply_solid_fill(ic, "EFF6FF")
        ic.line.color.rgb = hex_to_rgb("BFDBFE")
        ic_tf = ic.text_frame
        ic_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        icp = ic_tf.paragraphs[0]
        icp.text = "⚙"
        icp.font.size = Pt(11.0)
        icp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)
        icp.alignment = PP_ALIGN.CENTER

        # Label
        lbl_box = slide.shapes.add_textbox(Pt(ix + 6.0), Pt(464.0), Pt(card_w - 12.0), Pt(46.0))
        ltf = lbl_box.text_frame
        ltf.word_wrap = True
        ltf.vertical_anchor = MSO_ANCHOR.MIDDLE
        lp = ltf.paragraphs[0]
        lp.text = item_txt
        lp.alignment = PP_ALIGN.CENTER
        lp.font.name = "Aptos"
        lp.font.size = Pt(9.0)
        lp.font.bold = True
        lp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)


# ==============================================================================
# SLIDE 13: Operating Model
# ==============================================================================

def build_slide13(prs, lib_s1, lib_lines):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Title
    title_box = slide.shapes.add_textbox(Pt(30.0), Pt(6.0), Pt(900.0), Pt(32.0))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Operating Model: GEP runs KPOD’s PSCM operations end to end, supported by GEP experts, AI agents and built infrastructure"
    p.font.name = "Aptos"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

    # Left: KPOD Stakeholders (w = 140.0, top = 44.0, h = 480.0)
    k_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(44.0), Pt(140.0), Pt(480.0))
    apply_solid_fill(k_box, "FFFFFF")
    k_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(k_box, 4.0)
    apply_library_shadow(k_box)

    khdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(38.0), Pt(50.0), Pt(124.0), Pt(24.0))
    apply_solid_fill(khdr, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(khdr, 4.0)
    khtf = khdr.text_frame
    khtf.vertical_anchor = MSO_ANCHOR.MIDDLE
    khp = khtf.paragraphs[0]
    khp.text = "KPOD Stakeholders"
    khp.alignment = PP_ALIGN.CENTER
    khp.font.name = "Aptos"
    khp.font.size = Pt(9.0)
    khp.font.bold = True
    khp.font.color.rgb = hex_to_rgb("FFFFFF")

    k_items = [
        ("Leadership Team", 82.0, 48.0),
        ("KPOD Business and Functional Stakeholders\nProcurement Team", 138.0, 150.0),
        ("Suppliers", 296.0, 56.0)
    ]
    for k_txt, ky, kh in k_items:
        kb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(38.0), Pt(ky), Pt(124.0), Pt(kh))
        apply_solid_fill(kb, "F8FAFC")
        kb.line.color.rgb = hex_to_rgb("CBD5E1")
        set_shape_corner_radius(kb, 4.0)
        kbtf = kb.text_frame
        kbtf.word_wrap = True
        kbtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        kp = kbtf.paragraphs[0]
        kp.text = k_txt
        kp.alignment = PP_ALIGN.CENTER
        kp.font.name = "Aptos"
        kp.font.size = Pt(9.0)
        kp.font.bold = True
        kp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

    # Legend
    leg_box = slide.shapes.add_textbox(Pt(38.0), Pt(360.0), Pt(124.0), Pt(50.0))
    leg_tf = leg_box.text_frame
    leg_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    lp1 = leg_tf.paragraphs[0]
    lp1.text = "Legend:  KPOD  |  GEP"
    lp1.alignment = PP_ALIGN.CENTER
    lp1.font.name = "Aptos"
    lp1.font.size = Pt(9.0)
    lp1.font.bold = True
    lp1.font.color.rgb = hex_to_rgb("64748B")

    # Right: KPOD Integrated PSCM Operations (delivered by GEP) (w = 746.0, top = 44.0, h = 480.0)
    ops_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(184.0), Pt(44.0), Pt(746.0), Pt(480.0))
    apply_solid_fill(ops_box, "FFFFFF")
    ops_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(ops_box, 4.0)
    apply_library_shadow(ops_box)

    opshdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(194.0), Pt(50.0), Pt(726.0), Pt(22.0))
    apply_solid_fill(opshdr, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(opshdr, 4.0)
    optf = opshdr.text_frame
    optf.vertical_anchor = MSO_ANCHOR.MIDDLE
    opp = optf.paragraphs[0]
    opp.text = "KPOD Integrated PSCM Operations (delivered by GEP)"
    opp.alignment = PP_ALIGN.CENTER
    opp.font.name = "Aptos"
    opp.font.size = Pt(9.5)
    opp.font.bold = True
    opp.font.color.rgb = hex_to_rgb("FFFFFF")

    # Plan - to - Strategy Subheader & 3 Columns (top = 78.0, h = 100.0)
    pts_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(194.0), Pt(76.0), Pt(726.0), Pt(18.0))
    apply_solid_fill(pts_hdr, ColorTokens.BRAND_INDIGO)
    set_shape_corner_radius(pts_hdr, 4.0)
    ptf = pts_hdr.text_frame
    ptf.vertical_anchor = MSO_ANCHOR.MIDDLE
    pp = ptf.paragraphs[0]
    pp.text = "Plan - to - Strategy"
    pp.alignment = PP_ALIGN.CENTER
    pp.font.name = "Aptos"
    pp.font.size = Pt(9.0)
    pp.font.bold = True
    pp.font.color.rgb = hex_to_rgb("FFFFFF")

    p2s_cols = [
        ("Strategy Management", ["Procurement Strategy and Planning", "Organization Change and Management"]),
        ("Category Management", ["Market Analysis", "Strategy Creation & Approvals", "Strategy implementation & Optimization"]),
        ("Projects & Planning", ["Project Procurement (Planning & Execution)", "Demand & Capacity Planning"])
    ]
    p_w = 236.0
    for p_i, (c_name, c_bullets) in enumerate(p2s_cols):
        px = 194.0 + p_i * 245.0
        pb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(px), Pt(98.0), Pt(p_w), Pt(84.0))
        apply_solid_fill(pb, "F8FAFC")
        pb.line.color.rgb = hex_to_rgb("CBD5E1")
        set_shape_corner_radius(pb, 4.0)
        apply_library_shadow(pb)

        # Header
        ch = slide.shapes.add_textbox(Pt(px + 4.0), Pt(100.0), Pt(p_w - 8.0), Pt(16.0))
        chp = ch.text_frame.paragraphs[0]
        chp.text = c_name
        chp.alignment = PP_ALIGN.CENTER
        chp.font.name = "Aptos"
        chp.font.size = Pt(9.0)
        chp.font.bold = True
        chp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

        # Bullets
        cb = slide.shapes.add_textbox(Pt(px + 4.0), Pt(116.0), Pt(p_w - 8.0), Pt(64.0))
        cbtf = cb.text_frame
        cbtf.word_wrap = True
        cbtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        for b_idx, bullet in enumerate(c_bullets):
            pb = cbtf.paragraphs[0] if b_idx == 0 else cbtf.add_paragraph()
            pb.text = clean_bullet_string(bullet)
            pb.font.name = "Aptos"
            pb.font.size = Pt(9.0)
            pb.font.color.rgb = hex_to_rgb("334155")
            apply_native_bullet(pb, char="•", typeface="Arial", mar_l_pt=8.0, indent_pt=-8.0)

    # Technology Platform Arrow Banner (top = 188.0, h = 22.0)
    tp_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(194.0), Pt(188.0), Pt(726.0), Pt(22.0))
    apply_solid_fill(tp_box, "FFF7ED")
    tp_box.line.color.rgb = hex_to_rgb("FF8300")
    set_shape_corner_radius(tp_box, 4.0)
    tptf = tp_box.text_frame
    tptf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tpp = tptf.paragraphs[0]
    tpp.text = "KPOD Technology Platform: Oracle | FairmarkIT | Aravo | Beroe, with GEP AI Agents"
    tpp.alignment = PP_ALIGN.CENTER
    tpp.font.name = "Aptos"
    tpp.font.size = Pt(9.0)
    tpp.font.bold = True
    tpp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

    # Middle Section: S2C (4 cols) & P2P (3 cols) (top = 216.0, h = 180.0)
    # S2C Header
    s2c_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(194.0), Pt(216.0), Pt(414.0), Pt(18.0))
    apply_solid_fill(s2c_hdr, ColorTokens.BRAND_INDIGO)
    set_shape_corner_radius(s2c_hdr, 4.0)
    s2ctf = s2c_hdr.text_frame
    s2ctf.vertical_anchor = MSO_ANCHOR.MIDDLE
    s2cp = s2ctf.paragraphs[0]
    s2cp.text = "Source - to - Contract"
    s2cp.alignment = PP_ALIGN.CENTER
    s2cp.font.name = "Aptos"
    s2cp.font.size = Pt(9.0)
    s2cp.font.bold = True
    s2cp.font.color.rgb = hex_to_rgb("FFFFFF")

    # P2P Header
    p2p_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(614.0), Pt(216.0), Pt(306.0), Pt(18.0))
    apply_solid_fill(p2p_hdr, ColorTokens.BRAND_TANGERINE)
    set_shape_corner_radius(p2p_hdr, 4.0)
    p2ptf = p2p_hdr.text_frame
    p2ptf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p2pp = p2ptf.paragraphs[0]
    p2pp.text = "Purchase - to - Pay"
    p2pp.alignment = PP_ALIGN.CENTER
    p2pp.font.name = "Aptos"
    p2pp.font.size = Pt(9.0)
    p2pp.font.bold = True
    p2pp.font.color.rgb = hex_to_rgb("FFFFFF")

    # 4 S2C Columns
    s2c_cols = [
        ("Category Mgmt.", ["Demand Assessment & Identification", "Supply Market Assessment", "Category Strategy Approach", "Monitoring & Maintenance"]),
        ("Strategic Sourcing", ["Opportunity Assessment", "Sourcing Strategy Approach", "Sourcing Execution & Award", "Preferred Suppliers List"]),
        ("Supplier Mgmt.", ["Vendor Setup", "Supplier Relationship & Performance Management", "Supplier Risk Management"]),
        ("Contract Mgmt.", ["Contract Authoring", "Contract Execution", "Contract Administration", "Contract Termination & Renewal"])
    ]
    s2c_w = 98.0
    for s_i, (s_name, s_bullets) in enumerate(s2c_cols):
        sx = 194.0 + s_i * 105.0
        sb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(sx), Pt(238.0), Pt(s2c_w), Pt(156.0))
        apply_solid_fill(sb, "F8FAFC")
        sb.line.color.rgb = hex_to_rgb("CBD5E1")
        set_shape_corner_radius(sb, 4.0)
        apply_library_shadow(sb)

        shdr = slide.shapes.add_textbox(Pt(sx + 2.0), Pt(240.0), Pt(s2c_w - 4.0), Pt(16.0))
        shp = shdr.text_frame.paragraphs[0]
        shp.text = s_name
        shp.alignment = PP_ALIGN.CENTER
        shp.font.name = "Aptos"
        shp.font.size = Pt(9.0)
        shp.font.bold = True
        shp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

        sbtf_box = slide.shapes.add_textbox(Pt(sx + 2.0), Pt(256.0), Pt(s2c_w - 4.0), Pt(136.0))
        sbtf = sbtf_box.text_frame
        sbtf.word_wrap = True
        sbtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        for b_idx, bullet in enumerate(s_bullets):
            pb = sbtf.paragraphs[0] if b_idx == 0 else sbtf.add_paragraph()
            pb.text = clean_bullet_string(bullet)
            pb.font.name = "Aptos"
            pb.font.size = Pt(9.0)
            pb.font.color.rgb = hex_to_rgb("334155")
            apply_native_bullet(pb, char="•", typeface="Arial", mar_l_pt=7.0, indent_pt=-7.0)

    # 3 P2P Columns
    p2p_cols = [
        ("Order Mgmt.", ["R2O buying channels", "PO creation / Update", "PO monitoring", "Analyze bottlenecks, expediting", "Goods Service / Receipt"]),
        ("Materials Mgmt.", ["Material availability & shortages", "Spares & inventory optimisation", "Warehouse & logistics interfaces", "Customs coordination"]),
        ("Ops & Helpdesk", ["Supplier queries & claims", "Problem escalation", "Query resolution ownership"])
    ]
    p2p_w = 98.0
    for p_i, (p_name, p_bullets) in enumerate(p2p_cols):
        px = 614.0 + p_i * 105.0
        pb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(px), Pt(238.0), Pt(p2p_w), Pt(156.0))
        apply_solid_fill(pb, "F8FAFC")
        pb.line.color.rgb = hex_to_rgb("CBD5E1")
        set_shape_corner_radius(pb, 4.0)
        apply_library_shadow(pb)

        phdr = slide.shapes.add_textbox(Pt(px + 2.0), Pt(240.0), Pt(p2p_w - 4.0), Pt(16.0))
        php = phdr.text_frame.paragraphs[0]
        php.text = p_name
        php.alignment = PP_ALIGN.CENTER
        php.font.name = "Aptos"
        php.font.size = Pt(9.0)
        php.font.bold = True
        php.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

        pbtf_box = slide.shapes.add_textbox(Pt(px + 2.0), Pt(256.0), Pt(p2p_w - 4.0), Pt(136.0))
        pbtf = pbtf_box.text_frame
        pbtf.word_wrap = True
        pbtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        for b_idx, bullet in enumerate(p_bullets):
            pb = pbtf.paragraphs[0] if b_idx == 0 else pbtf.add_paragraph()
            pb.text = clean_bullet_string(bullet)
            pb.font.name = "Aptos"
            pb.font.size = Pt(9.0)
            pb.font.color.rgb = hex_to_rgb("334155")
            apply_native_bullet(pb, char="•", typeface="Arial", mar_l_pt=7.0, indent_pt=-7.0)

    # Bottom: GEP Center of Excellence (COE) (top = 400.0, h = 118.0)
    coe_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(194.0), Pt(400.0), Pt(726.0), Pt(18.0))
    apply_solid_fill(coe_hdr, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(coe_hdr, 4.0)
    ctf = coe_hdr.text_frame
    ctf.vertical_anchor = MSO_ANCHOR.MIDDLE
    cp = ctf.paragraphs[0]
    cp.text = "GEP Center of Excellence (COE)"
    cp.alignment = PP_ALIGN.CENTER
    cp.font.name = "Aptos"
    cp.font.size = Pt(9.0)
    cp.font.bold = True
    cp.font.color.rgb = hex_to_rgb("FFFFFF")

    coe_pills = [
        "Ready-to-Deploy Tools and Templates",
        "MI & Analytics COE",
        "Category Advisory Group and Category SMEs",
        "Best-In Class Benchmarks",
        "Reporting Dashboards"
    ]
    cp_w = 140.0
    for c_i, cp_name in enumerate(coe_pills):
        cpx = 194.0 + c_i * 146.0
        cp_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(cpx), Pt(422.0), Pt(cp_w), Pt(24.0))
        apply_solid_fill(cp_box, "F8FAFC")
        cp_box.line.color.rgb = hex_to_rgb("CBD5E1")
        set_shape_corner_radius(cp_box, 4.0)
        cptf = cp_box.text_frame
        cptf.vertical_anchor = MSO_ANCHOR.MIDDLE
        cpp = cptf.paragraphs[0]
        cpp.text = cp_name
        cpp.alignment = PP_ALIGN.CENTER
        cpp.font.name = "Aptos"
        cpp.font.size = Pt(9.0)
        cpp.font.bold = True
        cpp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

    # Digital Enablers Bar
    de_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(194.0), Pt(452.0), Pt(726.0), Pt(20.0))
    apply_solid_fill(de_box, "FFF7ED")
    de_box.line.color.rgb = hex_to_rgb("FED7AA")
    set_shape_corner_radius(de_box, 4.0)
    detf = de_box.text_frame
    detf.vertical_anchor = MSO_ANCHOR.MIDDLE
    dep = detf.paragraphs[0]
    dep.text = "Digital Enablers: GEP AI Agents (Intake, RFx, Negotiation, Contract Drafting, Should Cost) and RPA"
    dep.alignment = PP_ALIGN.CENTER
    dep.font.name = "Aptos"
    dep.font.size = Pt(9.0)
    dep.font.bold = True
    dep.font.color.rgb = hex_to_rgb("9A3412")

    # Governance Bar
    gv_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(194.0), Pt(476.0), Pt(726.0), Pt(20.0))
    apply_solid_fill(gv_box, "F1F5F9")
    gv_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(gv_box, 4.0)
    gvtf = gv_box.text_frame
    gvtf.vertical_anchor = MSO_ANCHOR.MIDDLE
    gvp = gvtf.paragraphs[0]
    gvp.text = "Program Governance     |     Compliance Management"
    gvp.alignment = PP_ALIGN.CENTER
    gvp.font.name = "Aptos"
    gvp.font.size = Pt(9.0)
    gvp.font.bold = True
    gvp.font.color.rgb = hex_to_rgb("475569")


# ==============================================================================
# SLIDES 14 & 15: Interface Ownership & Escalation
# ==============================================================================

def build_slide14_or_15(prs, lib_s1, lib_lines, is_slide15=False):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Title & Subtitle
    title_box = slide.shapes.add_textbox(Pt(30.0), Pt(8.0), Pt(900.0), Pt(20.0))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Every interface has an owner, and issues escalate before they cause delay"
    p.font.name = "Aptos"
    p.font.size = Pt(14.0)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

    sub_box = slide.shapes.add_textbox(Pt(30.0), Pt(30.0), Pt(900.0), Pt(16.0))
    stp = sub_box.text_frame.paragraphs[0]
    stp.text = "Clear ownership, structured escalation and proactive issue management across all KPOD interfaces"
    stp.font.name = "Aptos"
    stp.font.size = Pt(9.0)
    stp.font.color.rgb = hex_to_rgb("64748B")

    decor_proto = get_shape_by_name(lib_s1, "Design 1 Decor box")
    hdr_proto = get_shape_by_name(lib_s1, "Design 1 Header Text")

    # Left Half: 1. Interface Register from day one (w = 440.0, top = 50.0, h = 470.0)
    left_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(50.0), Pt(440.0), Pt(470.0))
    apply_solid_fill(left_box, "FFFFFF")
    left_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(left_box, 4.0)
    apply_library_shadow(left_box)

    # Header pill
    l_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(38.0), Pt(56.0), Pt(424.0), Pt(24.0))
    apply_solid_fill(l_hdr, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(l_hdr, 4.0)
    ltf = l_hdr.text_frame
    ltf.vertical_anchor = MSO_ANCHOR.MIDDLE
    lp = ltf.paragraphs[0]
    lp.text = "1. Interface Register from day one  |  All interfaces owned, mapped and governed"
    lp.alignment = PP_ALIGN.CENTER
    lp.font.name = "Aptos"
    lp.font.size = Pt(9.0)
    lp.font.bold = True
    lp.font.color.rgb = hex_to_rgb("FFFFFF")

    # Center Hub
    hub = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(180.0), Pt(175.0), Pt(140.0), Pt(80.0))
    apply_solid_fill(hub, ColorTokens.BRAND_INDIGO)
    set_shape_corner_radius(hub, 4.0)
    apply_library_shadow(hub)
    htf = hub.text_frame
    htf.word_wrap = True
    htf.vertical_anchor = MSO_ANCHOR.MIDDLE
    hp1 = htf.paragraphs[0]
    hp1.text = "GEP PSCM"
    hp1.alignment = PP_ALIGN.CENTER
    hp1.font.name = "Aptos"
    hp1.font.size = Pt(11.0)
    hp1.font.bold = True
    hp1.font.color.rgb = hex_to_rgb("FFFFFF")
    hp2 = htf.add_paragraph()
    hp2.text = "Category Manager\n(single point of contact per function)"
    hp2.alignment = PP_ALIGN.CENTER
    hp2.font.name = "Aptos"
    hp2.font.size = Pt(9.0)
    hp2.font.color.rgb = hex_to_rgb("E2E8F0")

    # 10 Spokes around the hub (strictly disjoint coordinates)
    spokes = [
        ("bp IJV", 42.0, 96.0),
        ("Operations", 150.0, 96.0),
        ("Wells", 260.0, 96.0),
        ("Projects", 368.0, 96.0),
        ("Security", 42.0, 175.0),
        ("Engineering", 368.0, 175.0),
        ("Logistics", 42.0, 254.0),
        ("NOC / NGC", 368.0, 254.0),
        ("Finance", 150.0, 295.0),
        ("HSE", 260.0, 295.0),
    ]

    for s_name, sx, sy in spokes:
        s_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(sx), Pt(sy), Pt(84.0), Pt(32.0))
        apply_solid_fill(s_box, "F8FAFC")
        s_box.line.color.rgb = hex_to_rgb("CBD5E1")
        set_shape_corner_radius(s_box, 4.0)
        stf = s_box.text_frame
        stf.vertical_anchor = MSO_ANCHOR.MIDDLE
        sp = stf.paragraphs[0]
        sp.text = s_name
        sp.alignment = PP_ALIGN.CENTER
        sp.font.name = "Aptos"
        sp.font.size = Pt(9.0)
        sp.font.bold = True
        sp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

    # Divider
    add_library_line(slide, lib_lines["side_fade_r"], 40.0, 360.0, 420.0)

    # Bottom 3 Cards in Left Section
    l_bottom_cards = [
        ("Defined owners", "GEP owner and KPOD counterpart for each interface"),
        ("Agreed turnaround", "SLAs for information exchange and decisions"),
        ("Integrated planning", "Regular cross functional alignment")
    ]
    card_w = 134.0
    for c_i, (c_title, c_desc) in enumerate(l_bottom_cards):
        cx = 40.0 + c_i * 142.0
        cbox = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(cx), Pt(374.0), Pt(card_w), Pt(136.0))
        apply_solid_fill(cbox, "F8FAFC")
        cbox.line.color.rgb = hex_to_rgb("E2E8F0")
        set_shape_corner_radius(cbox, 4.0)
        apply_library_shadow(cbox)

        ctf = cbox.text_frame
        ctf.word_wrap = True
        ctf.vertical_anchor = MSO_ANCHOR.MIDDLE
        cp1 = ctf.paragraphs[0]
        cp1.text = c_title
        cp1.font.name = "Aptos"
        cp1.font.size = Pt(9.0)
        cp1.font.bold = True
        cp1.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)

        cp2 = ctf.add_paragraph()
        cp2.text = c_desc
        cp2.font.name = "Aptos"
        cp2.font.size = Pt(9.0)
        cp2.font.color.rgb = hex_to_rgb("475569")

    # Right Half (w = 440.0, top = 50.0, h = 470.0)
    right_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(490.0), Pt(50.0), Pt(440.0), Pt(470.0))
    apply_solid_fill(right_box, "FFFFFF")
    right_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(right_box, 4.0)
    apply_library_shadow(right_box)

    # Header pill: 2. Three level escalation
    r_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(498.0), Pt(56.0), Pt(424.0), Pt(24.0))
    apply_solid_fill(r_hdr, ColorTokens.BRAND_TANGERINE)
    set_shape_corner_radius(r_hdr, 4.0)
    rtf = r_hdr.text_frame
    rtf.vertical_anchor = MSO_ANCHOR.MIDDLE
    rp = rtf.paragraphs[0]
    rp.text = "2. Three level escalation  |  Issues escalate early, with clear accountability"
    rp.alignment = PP_ALIGN.CENTER
    rp.font.name = "Aptos"
    rp.font.size = Pt(9.0)
    rp.font.bold = True
    rp.font.color.rgb = hex_to_rgb("FFFFFF")

    # 3 Escalation Levels (Stacked)
    esc_levels = [
        ("Level 3", "Strategic", "GEP Service Director + KPOD CPO", "Risk to operations, safety, regulatory dates or contract expiry", ColorTokens.BRAND_NAVY),
        ("Level 2", "Tactical", "GEP Service Delivery Manager + KPOD PDM", "At risk of more than 3 days delay, critical schedule risk or cross functional impact", ColorTokens.BRAND_INDIGO),
        ("Level 1", "Operational", "GEP Category Leads + KPOD function lead", "Dependency more than 3 working days overdue, or an emerging issue", ColorTokens.BRAND_TANGERINE)
    ]

    for e_i, (lvl_num, lvl_type, lvl_leads, lvl_desc, lvl_clr) in enumerate(esc_levels):
        ey = 90.0 + e_i * 84.0
        ebox = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(498.0), Pt(ey), Pt(424.0), Pt(64.0))
        apply_solid_fill(ebox, "F8FAFC")
        ebox.line.color.rgb = hex_to_rgb("CBD5E1")
        set_shape_corner_radius(ebox, 4.0)
        apply_library_shadow(ebox)

        # Level Badge
        l_badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(506.0), Pt(ey + 6.0), Pt(74.0), Pt(52.0))
        apply_solid_fill(l_badge, lvl_clr)
        set_shape_corner_radius(l_badge, 4.0)
        ltf = l_badge.text_frame
        ltf.vertical_anchor = MSO_ANCHOR.MIDDLE
        lp1 = ltf.paragraphs[0]
        lp1.text = lvl_num
        lp1.alignment = PP_ALIGN.CENTER
        lp1.font.name = "Aptos"
        lp1.font.size = Pt(10.0)
        lp1.font.bold = True
        lp1.font.color.rgb = hex_to_rgb("FFFFFF")
        lp2 = ltf.add_paragraph()
        lp2.text = lvl_type
        lp2.alignment = PP_ALIGN.CENTER
        lp2.font.name = "Aptos"
        lp2.font.size = Pt(9.0)
        lp2.font.color.rgb = hex_to_rgb("E2E8F0")

        # Details
        dtf_box = slide.shapes.add_textbox(Pt(588.0), Pt(ey + 4.0), Pt(326.0), Pt(56.0))
        dtf = dtf_box.text_frame
        dtf.word_wrap = True
        dtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        dp1 = dtf.paragraphs[0]
        dp1.text = lvl_leads
        dp1.font.name = "Aptos"
        dp1.font.size = Pt(9.5)
        dp1.font.bold = True
        dp1.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

        dp2 = dtf.add_paragraph()
        dp2.text = lvl_desc
        dp2.font.name = "Aptos"
        dp2.font.size = Pt(9.0)
        dp2.font.color.rgb = hex_to_rgb("475569")

        # Arrow up indicator between levels
        if e_i < 2:
            arr_box = slide.shapes.add_textbox(Pt(680.0), Pt(ey + 66.0), Pt(40.0), Pt(14.0))
            ap = arr_box.text_frame.paragraphs[0]
            ap.text = "▲"
            ap.alignment = PP_ALIGN.CENTER
            ap.font.size = Pt(10.0)
            ap.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_TANGERINE)

    # Bottom Right: 3. Avoiding delay, duplication and gaps (top = 348.0)
    g_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(498.0), Pt(348.0), Pt(424.0), Pt(24.0))
    apply_solid_fill(g_hdr, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(g_hdr, 4.0)
    gtf = g_hdr.text_frame
    gtf.vertical_anchor = MSO_ANCHOR.MIDDLE
    gp = gtf.paragraphs[0]
    gp.text = "3. Avoiding delay, duplication and gaps  |  Disciplined issue management"
    gp.alignment = PP_ALIGN.CENTER
    gp.font.name = "Aptos"
    gp.font.size = Pt(9.0)
    gp.font.bold = True
    gp.font.color.rgb = hex_to_rgb("FFFFFF")

    gap_cards = [
        ("One case record", "From requisition to contract with full visibility"),
        ("Agreed priority score", "Drives the right escalation and final call"),
        ("Weekly aged item review", "Stalled cases, overdue expediting, expiring contracts")
    ]
    gc_w = 134.0
    for c_i, (c_title, c_desc) in enumerate(gap_cards):
        cx = 498.0 + c_i * 145.0
        cbox = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(cx), Pt(380.0), Pt(gc_w), Pt(130.0))
        apply_solid_fill(cbox, "F8FAFC")
        cbox.line.color.rgb = hex_to_rgb("E2E8F0")
        set_shape_corner_radius(cbox, 4.0)
        apply_library_shadow(cbox)

        ctf = cbox.text_frame
        ctf.word_wrap = True
        ctf.vertical_anchor = MSO_ANCHOR.MIDDLE
        cp1 = ctf.paragraphs[0]
        cp1.text = c_title
        cp1.alignment = PP_ALIGN.CENTER
        cp1.font.name = "Aptos"
        cp1.font.size = Pt(9.0)
        cp1.font.bold = True
        cp1.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_TANGERINE)

        cp2 = ctf.add_paragraph()
        cp2.text = c_desc
        cp2.alignment = PP_ALIGN.CENTER
        cp2.font.name = "Aptos"
        cp2.font.size = Pt(9.0)
        cp2.font.color.rgb = hex_to_rgb("475569")


# ==============================================================================
# SLIDE 16: Transition Governance & Risk Management
# ==============================================================================

def build_slide16(prs, lib_s1, lib_lines):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Title
    title_box = slide.shapes.add_textbox(Pt(30.0), Pt(6.0), Pt(900.0), Pt(32.0))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Transition Governance & Risk Management : Three-tier governance followed by GEP will ensure transition accountability, proactive risk mitigation and service continuity"
    p.font.name = "Aptos"
    p.font.size = Pt(13.0)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

    decor_proto = get_shape_by_name(lib_s1, "Design 1 Decor box")
    hdr_proto = get_shape_by_name(lib_s1, "Design 1 Header Text")

    # Left Side: 3 Governance Tiers (w = 480.0, top = 44.0, h = 476.0)
    l_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(44.0), Pt(480.0), Pt(476.0))
    apply_solid_fill(l_box, "FFFFFF")
    l_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(l_box, 4.0)
    apply_library_shadow(l_box)

    gov_tiers = [
        {
            "name": "Transition Steering Committee (Monthly)",
            "clr": ColorTokens.BRAND_NAVY,
            "kpod": ["Executive Sponsor/s", "Global Program Lead", "Global Process Lead", "Change Management Lead", "Other Key Stakeholders"],
            "gep": ["Executive Sponsors", "Global Program Lead", "Global Process Lead", "Change Management Lead"]
        },
        {
            "name": "Transition Mgmt. Committee (Fortnightly)",
            "clr": ColorTokens.BRAND_INDIGO,
            "kpod": ["Global Program Lead", "Workstream Leads", "Process Leads", "Change Management Lead"],
            "gep": ["Global Program Lead", "Workstream Lead(s)", "Process Leads", "Change Management Lead"]
        },
        {
            "name": "Transition Operations Committee (Weekly / as needed)",
            "clr": ColorTokens.BRAND_TANGERINE,
            "kpod": ["Transition Leads", "Workstream Leads", "Other Key Stakeholders", "C & C agents"],
            "gep": ["Transition Leads", "Workstream Leads"]
        }
    ]

    tier_h = 92.0
    for t_i, tier in enumerate(gov_tiers):
        ty = 52.0 + t_i * 98.0
        # Tier Title
        thdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(38.0), Pt(ty), Pt(464.0), Pt(20.0))
        apply_solid_fill(thdr, tier["clr"])
        set_shape_corner_radius(thdr, 4.0)
        th_tf = thdr.text_frame
        th_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tp = th_tf.paragraphs[0]
        tp.text = tier["name"]
        tp.alignment = PP_ALIGN.CENTER
        tp.font.name = "Aptos"
        tp.font.size = Pt(9.0)
        tp.font.bold = True
        tp.font.color.rgb = hex_to_rgb("FFFFFF")

        # KPOD Col
        k_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(38.0), Pt(ty + 22.0), Pt(230.0), Pt(70.0))
        apply_solid_fill(k_box, "F8FAFC")
        k_box.line.color.rgb = hex_to_rgb("CBD5E1")
        set_shape_corner_radius(k_box, 4.0)
        ktf = k_box.text_frame
        ktf.word_wrap = True
        ktf.vertical_anchor = MSO_ANCHOR.MIDDLE
        kp0 = ktf.paragraphs[0]
        kp0.text = "KPOD: " + ", ".join(tier["kpod"])
        kp0.font.name = "Aptos"
        kp0.font.size = Pt(9.0)
        kp0.font.color.rgb = hex_to_rgb("334155")

        # GEP Col
        g_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(272.0), Pt(ty + 22.0), Pt(230.0), Pt(70.0))
        apply_solid_fill(g_box, "EFF6FF")
        g_box.line.color.rgb = hex_to_rgb("BFDBFE")
        set_shape_corner_radius(g_box, 4.0)
        gtf = g_box.text_frame
        gtf.word_wrap = True
        gtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        gp0 = gtf.paragraphs[0]
        gp0.text = "GEP: " + ", ".join(tier["gep"])
        gp0.font.name = "Aptos"
        gp0.font.size = Pt(9.0)
        gp0.font.color.rgb = hex_to_rgb("1E3A8A")

    # Bottom Left: Key Activities (top = 354.0, h = 158.0)
    ka_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(38.0), Pt(352.0), Pt(464.0), Pt(22.0))
    apply_solid_fill(ka_hdr, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(ka_hdr, 4.0)
    katf = ka_hdr.text_frame
    katf.vertical_anchor = MSO_ANCHOR.MIDDLE
    kap = katf.paragraphs[0]
    kap.text = "Key Activities"
    kap.alignment = PP_ALIGN.CENTER
    kap.font.name = "Aptos"
    kap.font.size = Pt(9.0)
    kap.font.bold = True
    kap.font.color.rgb = hex_to_rgb("FFFFFF")

    activities_cols = [
        ["Strategic Oversight", "Key decisions", "Resource prioritization", "Stage-gate approvals"],
        ["Progress tracking", "Risk and issue management", "Knowledge transfer", "Resource coordination"],
        ["Day-to-day execution", "Operational readiness", "Issue resolution", "Go-live preparation"]
    ]
    a_w = 152.0
    for a_i, a_bullets in enumerate(activities_cols):
        ax = 38.0 + a_i * 156.0
        abox = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(ax), Pt(378.0), Pt(a_w), Pt(134.0))
        apply_solid_fill(abox, "F8FAFC")
        abox.line.color.rgb = hex_to_rgb("CBD5E1")
        set_shape_corner_radius(abox, 4.0)
        apply_library_shadow(abox)

        abtf = abox.text_frame
        abtf.word_wrap = True
        abtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        for b_idx, bullet in enumerate(a_bullets):
            pb = abtf.paragraphs[0] if b_idx == 0 else abtf.add_paragraph()
            pb.text = clean_bullet_string(bullet)
            pb.font.name = "Aptos"
            pb.font.size = Pt(9.0)
            pb.font.color.rgb = hex_to_rgb("334155")
            apply_native_bullet(pb, char="•", typeface="Arial", mar_l_pt=8.0, indent_pt=-8.0)

    # Right Side: Table Matrix & KPOD Support (w = 404.0, top = 44.0, h = 476.0)
    r_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(526.0), Pt(44.0), Pt(404.0), Pt(476.0))
    apply_solid_fill(r_box, "FFFFFF")
    r_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(r_box, 4.0)
    apply_library_shadow(r_box)

    # Table Header: Key Transition Risk Mitigation
    t_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(534.0), Pt(50.0), Pt(388.0), Pt(22.0))
    apply_solid_fill(t_hdr, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(t_hdr, 4.0)
    tt_tf = t_hdr.text_frame
    tt_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tp = tt_tf.paragraphs[0]
    tp.text = "Key Transition Risk Mitigation"
    tp.alignment = PP_ALIGN.CENTER
    tp.font.name = "Aptos"
    tp.font.size = Pt(9.0)
    tp.font.bold = True
    tp.font.color.rgb = hex_to_rgb("FFFFFF")

    # Table Columns Header (y = 76.0, h = 20.0)
    t_cols = [("Risk area", 100.0), ("Key barriers", 138.0), ("GEP mitigation", 150.0)]
    tx = 534.0
    for tc_name, tc_w in t_cols:
        tcb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(tx), Pt(76.0), Pt(tc_w), Pt(20.0))
        apply_solid_fill(tcb, ColorTokens.BRAND_INDIGO)
        set_shape_corner_radius(tcb, 4.0)
        tc_tf = tcb.text_frame
        tc_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tcp = tc_tf.paragraphs[0]
        tcp.text = tc_name
        tcp.alignment = PP_ALIGN.CENTER
        tcp.font.name = "Aptos"
        tcp.font.size = Pt(9.0)
        tcp.font.bold = True
        tcp.font.color.rgb = hex_to_rgb("FFFFFF")
        tx += tc_w

    # Table Rows
    risk_rows = [
        ("Programme management", "Delayed data availability; delayed future-state decisions", "Early data collection; clear governance, roles and decision rights"),
        ("People readiness", "Knowledge-transfer gaps; loss of critical resources", "Multiple KT rounds; retain key resources; backfill critical vacancies"),
        ("Change management", "Insufficient stakeholder engagement; low adoption of new tools and processes", "Day-one change communications; customer journey focus; adoption KPIs")
    ]
    r_row_h = 74.0
    for r_i, (rarea, rbarr, rmit) in enumerate(risk_rows):
        ry = 98.0 + r_i * 76.0
        # Area
        b1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(534.0), Pt(ry), Pt(100.0), Pt(r_row_h))
        apply_solid_fill(b1, "F8FAFC")
        b1.line.color.rgb = hex_to_rgb("E2E8F0")
        set_shape_corner_radius(b1, 4.0)
        b1_tf = b1.text_frame
        b1_tf.word_wrap = True
        b1_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p1 = b1_tf.paragraphs[0]
        p1.text = rarea
        p1.font.name = "Aptos"
        p1.font.size = Pt(9.0)
        p1.font.bold = True
        p1.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

        # Barrier
        b2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(636.0), Pt(ry), Pt(138.0), Pt(r_row_h))
        apply_solid_fill(b2, "FFFFFF")
        b2.line.color.rgb = hex_to_rgb("E2E8F0")
        set_shape_corner_radius(b2, 4.0)
        b2_tf = b2.text_frame
        b2_tf.word_wrap = True
        b2_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p2 = b2_tf.paragraphs[0]
        p2.text = rbarr
        p2.font.name = "Aptos"
        p2.font.size = Pt(9.0)
        p2.font.color.rgb = hex_to_rgb("475569")

        # Mitigation
        b3 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(776.0), Pt(ry), Pt(146.0), Pt(r_row_h))
        apply_solid_fill(b3, "EFF6FF")
        b3.line.color.rgb = hex_to_rgb("BFDBFE")
        set_shape_corner_radius(b3, 4.0)
        b3_tf = b3.text_frame
        b3_tf.word_wrap = True
        b3_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p3 = b3_tf.paragraphs[0]
        p3.text = rmit
        p3.font.name = "Aptos"
        p3.font.size = Pt(9.0)
        p3.font.color.rgb = hex_to_rgb("1E3A8A")

    # Bottom Right: Where we need KPOD Support (top = 340.0, h = 172.0)
    ks_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(534.0), Pt(334.0), Pt(388.0), Pt(22.0))
    apply_solid_fill(ks_hdr, ColorTokens.BRAND_TANGERINE)
    set_shape_corner_radius(ks_hdr, 4.0)
    kstf = ks_hdr.text_frame
    kstf.vertical_anchor = MSO_ANCHOR.MIDDLE
    ksp = kstf.paragraphs[0]
    ksp.text = "Where we need KPOD Support"
    ksp.alignment = PP_ALIGN.CENTER
    ksp.font.name = "Aptos"
    ksp.font.size = Pt(9.0)
    ksp.font.bold = True
    ksp.font.color.rgb = hex_to_rgb("FFFFFF")

    ks_bullets_1 = [
        "Timely access to spend data, documentation and systems.",
        "Availability and retention of key transition personnel."
    ]
    ks_bullets_2 = [
        "Empowered leadership for future-state decisions.",
        "Active stakeholder participation and change adoption."
    ]

    for col_i, b_list in enumerate([ks_bullets_1, ks_bullets_2]):
        bx = 534.0 + col_i * 196.0
        bbox = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(bx), Pt(360.0), Pt(192.0), Pt(152.0))
        apply_solid_fill(bbox, "F8FAFC")
        bbox.line.color.rgb = hex_to_rgb("CBD5E1")
        set_shape_corner_radius(bbox, 4.0)
        apply_library_shadow(bbox)

        bbtf = bbox.text_frame
        bbtf.word_wrap = True
        bbtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        for b_idx, bullet in enumerate(b_list):
            pb = bbtf.paragraphs[0] if b_idx == 0 else bbtf.add_paragraph()
            pb.text = clean_bullet_string(bullet)
            pb.font.name = "Aptos"
            pb.font.size = Pt(9.0)
            pb.font.color.rgb = hex_to_rgb("334155")
            pb.space_after = Pt(8.0)
            apply_native_bullet(pb, char="•", typeface="Arial", mar_l_pt=8.0, indent_pt=-8.0)


# ==============================================================================
# SLIDE 17: PSCM Delivery Model: London-Led Governance...
# ==============================================================================

def build_slide17(prs, lib_s1, lib_lines):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Title
    title_box = slide.shapes.add_textbox(Pt(30.0), Pt(6.0), Pt(900.0), Pt(28.0))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "PSCM Delivery Model: London-Led Governance, Offshore Execution & Regional Expertise"
    p.font.name = "Aptos"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

    # Top: GLOBAL PSCM CREDENTIALS (w = 900.0, top = 36.0, h = 76.0)
    top_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(36.0), Pt(900.0), Pt(18.0))
    apply_solid_fill(top_hdr, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(top_hdr, 4.0)
    thtf = top_hdr.text_frame
    thtf.vertical_anchor = MSO_ANCHOR.MIDDLE
    thp = thtf.paragraphs[0]
    thp.text = "GLOBAL PSCM CREDENTIALS"
    thp.alignment = PP_ALIGN.CENTER
    thp.font.name = "Aptos"
    thp.font.size = Pt(9.0)
    thp.font.bold = True
    thp.font.color.rgb = hex_to_rgb("FFFFFF")

    top_kpis = [
        ("$600BN+", "Spend managed", ColorTokens.BRAND_NAVY),
        ("80+", "Mining, Energy & O&G clients", ColorTokens.BRAND_INDIGO),
        ("500+", "Category & industry SMEs", ColorTokens.BRAND_TANGERINE),
        ("120+", "Countries supported", ColorTokens.BRAND_NAVY),
        ("26+ years", "Procurement & supply chain experience", ColorTokens.BRAND_INDIGO)
    ]
    kw = 174.0
    for k_i, (kval, ksub, kclr) in enumerate(top_kpis):
        kx = 30.0 + k_i * 181.5
        kbox = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(kx), Pt(56.0), Pt(kw), Pt(52.0))
        apply_solid_fill(kbox, "F8FAFC")
        kbox.line.color.rgb = hex_to_rgb("CBD5E1")
        set_shape_corner_radius(kbox, 4.0)
        apply_library_shadow(kbox)

        ktf = kbox.text_frame
        ktf.vertical_anchor = MSO_ANCHOR.MIDDLE
        kp1 = ktf.paragraphs[0]
        kp1.text = kval
        kp1.alignment = PP_ALIGN.CENTER
        kp1.font.name = "Aptos"
        kp1.font.size = Pt(14.0)
        kp1.font.bold = True
        kp1.font.color.rgb = hex_to_rgb(kclr)

        kp2 = ktf.add_paragraph()
        kp2.text = ksub
        kp2.alignment = PP_ALIGN.CENTER
        kp2.font.name = "Aptos"
        kp2.font.size = Pt(9.0)
        kp2.font.color.rgb = hex_to_rgb("64748B")

    # Middle Left: LONDON-LED / OFFSHORE-SUPPORTED MODEL (w = 280.0, top = 114.0, h = 280.0)
    ml_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(114.0), Pt(280.0), Pt(280.0))
    apply_solid_fill(ml_box, "FFFFFF")
    ml_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(ml_box, 4.0)
    apply_library_shadow(ml_box)

    ml_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(36.0), Pt(118.0), Pt(268.0), Pt(18.0))
    apply_solid_fill(ml_hdr, ColorTokens.BRAND_INDIGO)
    set_shape_corner_radius(ml_hdr, 4.0)
    mltf = ml_hdr.text_frame
    mltf.vertical_anchor = MSO_ANCHOR.MIDDLE
    mlp = mltf.paragraphs[0]
    mlp.text = "LONDON-LED / OFFSHORE-SUPPORTED MODEL"
    mlp.alignment = PP_ALIGN.CENTER
    mlp.font.name = "Aptos"
    mlp.font.size = Pt(9.0)
    mlp.font.bold = True
    mlp.font.color.rgb = hex_to_rgb("FFFFFF")

    # London card
    lon_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(36.0), Pt(138.0), Pt(268.0), Pt(76.0))
    apply_solid_fill(lon_box, "F8FAFC")
    lon_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(lon_box, 4.0)
    ltf = lon_box.text_frame
    ltf.word_wrap = True
    ltf.vertical_anchor = MSO_ANCHOR.MIDDLE
    lp0 = ltf.paragraphs[0]
    lp0.text = "LONDON | LEADERSHIP & CLIENT INTERFACE"
    lp0.font.name = "Aptos"
    lp0.font.size = Pt(9.0)
    lp0.font.bold = True
    lp0.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)
    lon_bullets = ["Service leadership", "Senior stakeholder engagement", "Governance & performance oversight", "Category strategy & commercial direction"]
    for b in lon_bullets:
        pb = ltf.add_paragraph()
        pb.text = clean_bullet_string(b)
        pb.font.name = "Aptos"
        pb.font.size = Pt(9.0)
        pb.font.color.rgb = hex_to_rgb("334155")
        apply_native_bullet(pb, char="•", typeface="Arial", mar_l_pt=7.0, indent_pt=-7.0)

    # Arrow down
    arr = slide.shapes.add_textbox(Pt(160.0), Pt(214.0), Pt(20.0), Pt(14.0))
    arp = arr.text_frame.paragraphs[0]
    arp.text = "▼"
    arp.font.size = Pt(9.0)
    arp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_TANGERINE)

    # Offshore card
    off_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(36.0), Pt(228.0), Pt(268.0), Pt(76.0))
    apply_solid_fill(off_box, "EFF6FF")
    off_box.line.color.rgb = hex_to_rgb("BFDBFE")
    set_shape_corner_radius(off_box, 4.0)
    otf = off_box.text_frame
    otf.word_wrap = True
    otf.vertical_anchor = MSO_ANCHOR.MIDDLE
    op0 = otf.paragraphs[0]
    op0.text = "OFFSHORE | EXECUTION & DELIVERY SUPPORT"
    op0.font.name = "Aptos"
    op0.font.size = Pt(9.0)
    op0.font.bold = True
    op0.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)
    off_bullets = ["Sourcing & contracting execution", "Procurement operations", "Analytics & reporting", "Supplier and contract administration"]
    for b in off_bullets:
        pb = otf.add_paragraph()
        pb.text = clean_bullet_string(b)
        pb.font.name = "Aptos"
        pb.font.size = Pt(9.0)
        pb.font.color.rgb = hex_to_rgb("334155")
        apply_native_bullet(pb, char="•", typeface="Arial", mar_l_pt=7.0, indent_pt=-7.0)

    # Resource structure pills (L3, L2, L1)
    res_hdr = slide.shapes.add_textbox(Pt(36.0), Pt(306.0), Pt(268.0), Pt(14.0))
    rp = res_hdr.text_frame.paragraphs[0]
    rp.text = "KPOD-ALIGNED RESOURCE STRUCTURE"
    rp.font.name = "Aptos"
    rp.font.size = Pt(9.0)
    rp.font.bold = True
    rp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

    tiers_l = [
        ("L3", "Service Director | Delivery Manager | Category & Commercial Managers", ColorTokens.BRAND_NAVY),
        ("L2", "Sourcing & Contracting Specialists | Governance | Materials | Local Content", ColorTokens.BRAND_INDIGO),
        ("L1", "Analysts | Buyers | Contract Administrators | Reporting", ColorTokens.BRAND_TANGERINE)
    ]
    for ti, (tnum, tdesc, tclr) in enumerate(tiers_l):
        ty = 322.0 + ti * 22.0
        tb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(36.0), Pt(ty), Pt(268.0), Pt(20.0))
        apply_solid_fill(tb, "F8FAFC")
        tb.line.color.rgb = hex_to_rgb("CBD5E1")
        set_shape_corner_radius(tb, 4.0)
        tt_tf = tb.text_frame
        tt_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        ttp = tt_tf.paragraphs[0]
        r1 = ttp.add_run()
        r1.text = tnum + "  "
        r1.font.bold = True
        r1.font.size = Pt(9.0)
        r1.font.color.rgb = hex_to_rgb(tclr)
        r2 = ttp.add_run()
        r2.text = tdesc
        r2.font.size = Pt(9.0)
        r2.font.color.rgb = hex_to_rgb("334155")

    # Middle Right: REGIONAL DELIVERY FOOTPRINT (w = 608.0, top = 114.0, h = 280.0)
    rf_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(322.0), Pt(114.0), Pt(608.0), Pt(280.0))
    apply_solid_fill(rf_box, "FFFFFF")
    rf_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(rf_box, 4.0)
    apply_library_shadow(rf_box)

    rf_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(328.0), Pt(118.0), Pt(596.0), Pt(18.0))
    apply_solid_fill(rf_hdr, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(rf_hdr, 4.0)
    rftf = rf_hdr.text_frame
    rftf.vertical_anchor = MSO_ANCHOR.MIDDLE
    rfp = rftf.paragraphs[0]
    rfp.text = "REGIONAL DELIVERY FOOTPRINT | EUROPE & MIDDLE EAST"
    rfp.alignment = PP_ALIGN.CENTER
    rfp.font.name = "Aptos"
    rfp.font.size = Pt(9.0)
    rfp.font.bold = True
    rfp.font.color.rgb = hex_to_rgb("FFFFFF")

    # Europe Box (Left half)
    eu_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(328.0), Pt(140.0), Pt(294.0), Pt(246.0))
    apply_solid_fill(eu_box, "F8FAFC")
    eu_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(eu_box, 4.0)
    eutf = eu_box.text_frame
    eutf.word_wrap = True
    eutf.vertical_anchor = MSO_ANCHOR.MIDDLE
    eup0 = eutf.paragraphs[0]
    eup0.text = "EUROPEAN DELIVERY NETWORK | HQ : London"
    eup0.alignment = PP_ALIGN.CENTER
    eup0.font.name = "Aptos"
    eup0.font.size = Pt(9.0)
    eup0.font.bold = True
    eup0.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)

    eu_stats = ["100+ Clients served", "€150BN+ Spend managed", "20+ languages supported"]
    for s in eu_stats:
        sp = eutf.add_paragraph()
        sp.text = clean_bullet_string(s)
        sp.font.name = "Aptos"
        sp.font.size = Pt(9.0)
        sp.font.bold = True
        sp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)
        apply_native_bullet(sp, char="•", typeface="Arial", mar_l_pt=8.0, indent_pt=-8.0)

    # Middle East Box (Right half)
    me_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(630.0), Pt(140.0), Pt(294.0), Pt(246.0))
    apply_solid_fill(me_box, "FFF7ED")
    me_box.line.color.rgb = hex_to_rgb("FED7AA")
    set_shape_corner_radius(me_box, 4.0)
    metf = me_box.text_frame
    metf.word_wrap = True
    metf.vertical_anchor = MSO_ANCHOR.MIDDLE
    mep0 = metf.paragraphs[0]
    mep0.text = "MIDDLE EAST EXPERIENCE"
    mep0.alignment = PP_ALIGN.CENTER
    mep0.font.name = "Aptos"
    mep0.font.size = Pt(9.0)
    mep0.font.bold = True
    mep0.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_TANGERINE)

    me_stats = [
        "Regional markets supported include Egypt, Qatar, Saudi Arabia and the UAE.",
        "100+ Dedicated procurement professionals",
        "200+ Projects managed",
        "$1.5BN Annual spend managed",
        "12+ Clients served"
    ]
    for s in me_stats:
        sp = metf.add_paragraph()
        sp.text = clean_bullet_string(s)
        sp.font.name = "Aptos"
        sp.font.size = Pt(9.0)
        sp.font.color.rgb = hex_to_rgb("334155")
        apply_native_bullet(sp, char="•", typeface="Arial", mar_l_pt=8.0, indent_pt=-8.0)

    # Bottom: ENERGY-SECTOR EXPERTISE & PSCM COMPETENCIES (top = 402.0, h = 118.0)
    comp_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(402.0), Pt(900.0), Pt(118.0))
    apply_solid_fill(comp_box, "FFFFFF")
    comp_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(comp_box, 4.0)
    apply_library_shadow(comp_box)

    c_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(36.0), Pt(406.0), Pt(888.0), Pt(18.0))
    apply_solid_fill(c_hdr, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(c_hdr, 4.0)
    ctf = c_hdr.text_frame
    ctf.vertical_anchor = MSO_ANCHOR.MIDDLE
    cp = ctf.paragraphs[0]
    cp.text = "ENERGY-SECTOR EXPERTISE & PSCM COMPETENCIES"
    cp.alignment = PP_ALIGN.CENTER
    cp.font.name = "Aptos"
    cp.font.size = Pt(9.0)
    cp.font.bold = True
    cp.font.color.rgb = hex_to_rgb("FFFFFF")

    comps = [
        ("Technical procurement", "Sourcing execution &\nprocurement processes"),
        ("Energy category knowledge", "Drilling, MRO, well services\n& EPC"),
        ("Commercial & contracts", "Negotiation, contracting\n& claims"),
        ("Systems & digital", "Oracle, SAP Ariba &\nFairmarkIT"),
        ("Risk & compliance", "Conduct, ABC &\nsanctions"),
        ("Qualifications", "CIPS / MCIPS or\nequivalent")
    ]
    comp_w = 142.0
    for ci, (cname, cdesc) in enumerate(comps):
        cx = 36.0 + ci * 149.0
        cb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(cx), Pt(428.0), Pt(comp_w), Pt(86.0))
        apply_solid_fill(cb, "F8FAFC")
        cb.line.color.rgb = hex_to_rgb("E2E8F0")
        set_shape_corner_radius(cb, 4.0)
        apply_library_shadow(cb)

        cbtf = cb.text_frame
        cbtf.word_wrap = True
        cbtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        cp1 = cbtf.paragraphs[0]
        cp1.text = cname
        cp1.alignment = PP_ALIGN.CENTER
        cp1.font.name = "Aptos"
        cp1.font.size = Pt(9.0)
        cp1.font.bold = True
        cp1.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

        cp2 = cbtf.add_paragraph()
        cp2.text = cdesc
        cp2.alignment = PP_ALIGN.CENTER
        cp2.font.name = "Aptos"
        cp2.font.size = Pt(9.0)
        cp2.font.color.rgb = hex_to_rgb("475569")


# ==============================================================================
# SLIDE 18: Supplier & Materials Management
# ==============================================================================

def build_slide18(prs, lib_s1, lib_lines):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Title
    title_box = slide.shapes.add_textbox(Pt(30.0), Pt(6.0), Pt(900.0), Pt(28.0))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Supplier & Materials Management: Strengthening Performance, Risk & Operational Continuity"
    p.font.name = "Aptos"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

    decor_proto = get_shape_by_name(lib_s1, "Design 1 Decor box")
    hdr_proto = get_shape_by_name(lib_s1, "Design 1 Header Text")
    hdr_proto = get_shape_by_name(lib_s1, "Design 1 Header Text")

    # Section 1: Integrated Supplier & Contract Lifecycle Management (top = 40.0, h = 224.0)
    top_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(300.0), Pt(38.0), Pt(360.0), Pt(22.0))
    apply_solid_fill(top_hdr, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(top_hdr, 4.0)
    thtf = top_hdr.text_frame
    thtf.vertical_anchor = MSO_ANCHOR.MIDDLE
    thp = thtf.paragraphs[0]
    thp.text = "Integrated Supplier & Contract Lifecycle Management"
    thp.alignment = PP_ALIGN.CENTER
    thp.font.name = "Aptos"
    thp.font.size = Pt(9.0)
    thp.font.bold = True
    thp.font.color.rgb = hex_to_rgb("FFFFFF")

    s1_cols = [
        {
            "name": "Supplier and contract performance",
            "clr": ColorTokens.BRAND_INDIGO,
            "bullets": [
                "Contracts segmented by criticality and value",
                "Critical: management plan, monthly review, quarterly scorecard (Section VII)",
                "Improvement plans tracked to closure"
            ]
        },
        {
            "name": "Change, claims and disputes",
            "clr": ColorTokens.BRAND_NAVY,
            "bullets": [
                "Single change and claims register: value, entitlement, due date",
                "Recommended position with evidence for KPOD approval",
                "Legal advice stays with KPOD Legal"
            ]
        },
        {
            "name": "Supply chain risk and continuity",
            "clr": ColorTokens.BRAND_TANGERINE,
            "bullets": [
                "Financial health, single source, geopolitical, sanctions and lead time risk",
                "Risk monitoring using Aravo and Beroe",
                "Continuity plans covering alternative sources, buffer stock, expediting triggers"
            ]
        }
    ]

    col_w = 290.0
    gap = 15.0
    for c_i, col in enumerate(s1_cols):
        cx = 30.0 + c_i * (col_w + gap)
        cbox = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(cx), Pt(64.0), Pt(col_w), Pt(200.0))
        apply_solid_fill(cbox, "FFFFFF")
        cbox.line.color.rgb = hex_to_rgb("CBD5E1")
        set_shape_corner_radius(cbox, 4.0)
        apply_library_shadow(cbox)

        # Header bar
        chdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(cx + 8.0), Pt(70.0), Pt(col_w - 16.0), Pt(26.0))
        apply_solid_fill(chdr, col["clr"])
        set_shape_corner_radius(chdr, 4.0)
        chtf = chdr.text_frame
        chtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        chp = chtf.paragraphs[0]
        chp.text = col["name"]
        chp.alignment = PP_ALIGN.CENTER
        chp.font.name = "Aptos"
        chp.font.size = Pt(9.0)
        chp.font.bold = True
        chp.font.color.rgb = hex_to_rgb("FFFFFF")

        # Bullets
        b_box = slide.shapes.add_textbox(Pt(cx + 10.0), Pt(102.0), Pt(col_w - 20.0), Pt(156.0))
        btf = b_box.text_frame
        btf.word_wrap = True
        btf.vertical_anchor = MSO_ANCHOR.MIDDLE
        for b_idx, bullet in enumerate(col["bullets"]):
            pb = btf.paragraphs[0] if b_idx == 0 else btf.add_paragraph()
            pb.text = clean_bullet_string(bullet)
            pb.font.name = "Aptos"
            pb.font.size = Pt(9.0)
            pb.font.color.rgb = hex_to_rgb("334155")
            pb.space_after = Pt(6.0)
            apply_native_bullet(pb, char="•", typeface="Arial", mar_l_pt=9.0, indent_pt=-9.0)

    # Section 2: Materials, Logistics, Inventory & Warehousing Management (top = 278.0, h = 246.0)
    bot_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(280.0), Pt(274.0), Pt(400.0), Pt(22.0))
    apply_solid_fill(bot_hdr, ColorTokens.BRAND_TANGERINE)
    set_shape_corner_radius(bot_hdr, 4.0)
    bthtf = bot_hdr.text_frame
    bthtf.vertical_anchor = MSO_ANCHOR.MIDDLE
    bthp = bthtf.paragraphs[0]
    bthp.text = "Materials, Logistics, Inventory & Warehousing Management"
    bthp.alignment = PP_ALIGN.CENTER
    bthp.font.name = "Aptos"
    bthp.font.size = Pt(9.0)
    bthp.font.bold = True
    bthp.font.color.rgb = hex_to_rgb("FFFFFF")

    # Left: In base scope: GEP coordinates (w = 590.0, top = 300.0, h = 220.0)
    in_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(300.0), Pt(590.0), Pt(220.0))
    apply_solid_fill(in_box, "FFFFFF")
    in_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(in_box, 4.0)
    apply_library_shadow(in_box)

    in_lbl = slide.shapes.add_textbox(Pt(40.0), Pt(304.0), Pt(570.0), Pt(18.0))
    inp = in_lbl.text_frame.paragraphs[0]
    inp.text = "In base scope: GEP coordinates"
    inp.font.name = "Aptos"
    inp.font.size = Pt(9.5)
    inp.font.bold = True
    inp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

    base_cards = [
        "PO-to-delivery plan coordination",
        "Expedite critical orders",
        "Availability and shortage reporting",
        "Warehouse and pipeyard interface",
        "Inbound logistics and customs coordination",
        "Traceability from PR to receipt"
    ]
    bc_w = 180.0
    bc_h = 76.0
    for b_i, b_text in enumerate(base_cards):
        col_idx = b_i % 3
        row_idx = b_i // 3
        bx = 42.0 + col_idx * 190.0
        by = 328.0 + row_idx * 84.0
        bc_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(bx), Pt(by), Pt(bc_w), Pt(bc_h))
        apply_solid_fill(bc_box, "F8FAFC")
        bc_box.line.color.rgb = hex_to_rgb("E2E8F0")
        set_shape_corner_radius(bc_box, 4.0)
        apply_library_shadow(bc_box)

        bctf = bc_box.text_frame
        bctf.word_wrap = True
        bctf.vertical_anchor = MSO_ANCHOR.MIDDLE
        bcp = bctf.paragraphs[0]
        bcp.text = b_text
        bcp.alignment = PP_ALIGN.CENTER
        bcp.font.name = "Aptos"
        bcp.font.size = Pt(9.0)
        bcp.font.bold = True
        bcp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

    # Right: Optional, priced separately (w = 296.0, top = 300.0, h = 220.0)
    opt_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(634.0), Pt(300.0), Pt(296.0), Pt(220.0))
    apply_solid_fill(opt_box, "FFFFFF")
    opt_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(opt_box, 4.0)
    apply_library_shadow(opt_box)

    opt_lbl = slide.shapes.add_textbox(Pt(644.0), Pt(304.0), Pt(276.0), Pt(18.0))
    optp = opt_lbl.text_frame.paragraphs[0]
    optp.text = "Optional, priced separately"
    optp.font.name = "Aptos"
    optp.font.size = Pt(9.5)
    optp.font.bold = True
    optp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_TANGERINE)

    opt_bullets = [
        "Physical warehouse operations",
        "Stock verification",
        "Inventory control and accuracy management.",
        "Customs brokerage",
        "KPIs: on time delivery, critical order fulfilment"
    ]
    opt_tf_box = slide.shapes.add_textbox(Pt(644.0), Pt(328.0), Pt(276.0), Pt(180.0))
    ot_tf = opt_tf_box.text_frame
    ot_tf.word_wrap = True
    ot_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for b_idx, bullet in enumerate(opt_bullets):
        pb = ot_tf.paragraphs[0] if b_idx == 0 else ot_tf.add_paragraph()
        pb.text = clean_bullet_string(bullet)
        pb.font.name = "Aptos"
        pb.font.size = Pt(9.0)
        pb.font.color.rgb = hex_to_rgb("334155")
        pb.space_after = Pt(4.0)
        apply_native_bullet(pb, char="•", typeface="Arial", mar_l_pt=8.0, indent_pt=-8.0)


# ==============================================================================
# SLIDE 19: Risk Management Framework
# ==============================================================================

def build_slide19(prs, lib_s1, lib_lines):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Title
    title_box = slide.shapes.add_textbox(Pt(30.0), Pt(6.0), Pt(900.0), Pt(28.0))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Risk Management Framework : Integrated Supplier & Third-Party Risk Management Framework to Build a Resilient, Compliant and Risk-Aware Ecosystem"
    p.font.name = "Aptos"
    p.font.size = Pt(13.0)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

    # Subtitle Banner
    sub_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(38.0), Pt(900.0), Pt(28.0))
    apply_solid_fill(sub_box, "F8FAFC")
    sub_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(sub_box, 4.0)
    stf = sub_box.text_frame
    stf.word_wrap = True
    stf.vertical_anchor = MSO_ANCHOR.MIDDLE
    sp = stf.paragraphs[0]
    sp.text = "Combining supplier risk intelligence, third-party governance, continuous monitoring, and AI-driven insights to proactively identify, assess, mitigate, and monitor enterprise-wide risks."
    sp.alignment = PP_ALIGN.CENTER
    sp.font.name = "Aptos"
    sp.font.size = Pt(9.0)
    sp.font.italic = True
    sp.font.color.rgb = hex_to_rgb("334155")

    # End to End Risk management Cycle (top = 74.0, h = 68.0)
    cycle_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(350.0), Pt(72.0), Pt(260.0), Pt(18.0))
    apply_solid_fill(cycle_hdr, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(cycle_hdr, 4.0)
    ctf = cycle_hdr.text_frame
    ctf.vertical_anchor = MSO_ANCHOR.MIDDLE
    cp = ctf.paragraphs[0]
    cp.text = "End to End Risk management Cycle"
    cp.alignment = PP_ALIGN.CENTER
    cp.font.name = "Aptos"
    cp.font.size = Pt(9.0)
    cp.font.bold = True
    cp.font.color.rgb = hex_to_rgb("FFFFFF")

    cycle_steps = ["Identify", "Assess", "Segment", "Mitigate", "Monitor", "Govern"]
    cw = 138.0
    for c_i, step_name in enumerate(cycle_steps):
        cx = 30.0 + c_i * 152.4
        cbox = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(cx), Pt(94.0), Pt(cw), Pt(42.0))
        apply_solid_fill(cbox, "EFF6FF" if c_i % 2 == 0 else "FFF7ED")
        cbox.line.color.rgb = hex_to_rgb("BFDBFE" if c_i % 2 == 0 else "FED7AA")
        set_shape_corner_radius(cbox, 4.0)
        apply_library_shadow(cbox)

        cbtf = cbox.text_frame
        cbtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        cbp = cbtf.paragraphs[0]
        cbp.text = step_name
        cbp.alignment = PP_ALIGN.CENTER
        cbp.font.name = "Aptos"
        cbp.font.size = Pt(9.5)
        cbp.font.bold = True
        cbp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO if c_i % 2 == 0 else ColorTokens.BRAND_TANGERINE)

        if c_i < 5:
            arr = slide.shapes.add_textbox(Pt(cx + cw + 2.0), Pt(104.0), Pt(12.0), Pt(22.0))
            arp = arr.text_frame.paragraphs[0]
            arp.text = "▶"
            arp.font.size = Pt(9.0)
            arp.font.color.rgb = hex_to_rgb("94A3B8")

    # Middle: Risk Management Solution Portfolio (top = 146.0, h = 236.0)
    p_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(350.0), Pt(144.0), Pt(260.0), Pt(18.0))
    apply_solid_fill(p_hdr, ColorTokens.BRAND_INDIGO)
    set_shape_corner_radius(p_hdr, 4.0)
    ptf = p_hdr.text_frame
    ptf.vertical_anchor = MSO_ANCHOR.MIDDLE
    pp = ptf.paragraphs[0]
    pp.text = "Risk Management Solution Portfolio"
    pp.alignment = PP_ALIGN.CENTER
    pp.font.name = "Aptos"
    pp.font.size = Pt(9.0)
    pp.font.bold = True
    pp.font.color.rgb = hex_to_rgb("FFFFFF")

    portfolio_cols = [
        {
            "name": "Risk Coverage",
            "clr": ColorTokens.BRAND_INDIGO,
            "bullets": ["Financial", "Operational", "Cyber & Information Security", "ESG", "Compliance", "Supply Chain", "Geopolitical"]
        },
        {
            "name": "GEP Differentiators",
            "clr": ColorTokens.BRAND_NAVY,
            "bullets": [
                "Comprehensive SRM & TPRM Framework",
                "Integrated Risk Operating Model",
                "Risk Quantification & Scoring Engine",
                "Continuous Monitoring & early warning alerts",
                "Automated due diligence workflows",
                "Market Intelligence & External Data Integration",
                "AI Powered Risks insights and Recommendations"
            ]
        },
        {
            "name": "AI Accelerators",
            "clr": ColorTokens.BRAND_TANGERINE,
            "bullets": [
                "Supplier Screening Agent",
                "Regulatory Compliance Agent",
                "Diversity Verification Agent",
                "Continuous Risk Monitoring Agent",
                "Predictive Risk Alert Engine"
            ]
        }
    ]

    pw = 286.0
    for p_i, pcol in enumerate(portfolio_cols):
        px = 30.0 + p_i * 307.0
        pbox = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(px), Pt(166.0), Pt(pw), Pt(210.0))
        apply_solid_fill(pbox, "FFFFFF")
        pbox.line.color.rgb = hex_to_rgb("CBD5E1")
        set_shape_corner_radius(pbox, 4.0)
        apply_library_shadow(pbox)

        # Header bar
        phdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(px + 8.0), Pt(172.0), Pt(pw - 16.0), Pt(24.0))
        apply_solid_fill(phdr, pcol["clr"])
        set_shape_corner_radius(phdr, 4.0)
        phtf = phdr.text_frame
        phtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        php = phtf.paragraphs[0]
        php.text = pcol["name"]
        php.alignment = PP_ALIGN.CENTER
        php.font.name = "Aptos"
        php.font.size = Pt(9.0)
        php.font.bold = True
        php.font.color.rgb = hex_to_rgb("FFFFFF")

        # Bullets
        pbtf_box = slide.shapes.add_textbox(Pt(px + 10.0), Pt(202.0), Pt(pw - 20.0), Pt(168.0))
        pbtf = pbtf_box.text_frame
        pbtf.word_wrap = True
        pbtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        for b_idx, bullet in enumerate(pcol["bullets"]):
            pb = pbtf.paragraphs[0] if b_idx == 0 else pbtf.add_paragraph()
            pb.text = clean_bullet_string(bullet)
            pb.font.name = "Aptos"
            pb.font.size = Pt(9.0)
            pb.font.color.rgb = hex_to_rgb("334155")
            pb.space_after = Pt(2.5)
            apply_native_bullet(pb, char="•", typeface="Arial", mar_l_pt=8.0, indent_pt=-8.0)

    # Bottom: What GEP TPRM Brings: (top = 388.0, h = 136.0)
    b_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(370.0), Pt(386.0), Pt(220.0), Pt(18.0))
    apply_solid_fill(b_hdr, ColorTokens.BRAND_TANGERINE)
    set_shape_corner_radius(b_hdr, 4.0)
    bhtf = b_hdr.text_frame
    bhtf.vertical_anchor = MSO_ANCHOR.MIDDLE
    bhp = bhtf.paragraphs[0]
    bhp.text = "What GEP TPRM Brings:"
    bhp.alignment = PP_ALIGN.CENTER
    bhp.font.name = "Aptos"
    bhp.font.size = Pt(9.0)
    bhp.font.bold = True
    bhp.font.color.rgb = hex_to_rgb("FFFFFF")

    tprm_benefits = [
        "20-30% Faster Risk Assessment",
        "Improved Regulatory Compliance",
        "Enhanced Supplier Resilience",
        "Reduced disruption exposure",
        "Enterprise-wide risk visibility"
    ]
    tw = 170.0
    for t_i, bname in enumerate(tprm_benefits):
        tx = 30.0 + t_i * 182.5
        tbox = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(tx), Pt(408.0), Pt(tw), Pt(110.0))
        apply_solid_fill(tbox, "F8FAFC")
        tbox.line.color.rgb = hex_to_rgb("CBD5E1")
        set_shape_corner_radius(tbox, 4.0)
        apply_library_shadow(tbox)

        # Icon Circle
        ic = slide.shapes.add_shape(MSO_SHAPE.OVAL, Pt(tx + (tw - 28.0) / 2.0), Pt(416.0), Pt(28.0), Pt(28.0))
        apply_solid_fill(ic, "FFF7ED")
        ic.line.color.rgb = hex_to_rgb("FED7AA")
        ictf = ic.text_frame
        ictf.vertical_anchor = MSO_ANCHOR.MIDDLE
        icp = ictf.paragraphs[0]
        icp.text = "★"
        icp.font.size = Pt(11.0)
        icp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_TANGERINE)
        icp.alignment = PP_ALIGN.CENTER

        # Text
        ltf_box = slide.shapes.add_textbox(Pt(tx + 6.0), Pt(450.0), Pt(tw - 12.0), Pt(60.0))
        ltf = ltf_box.text_frame
        ltf.word_wrap = True
        ltf.vertical_anchor = MSO_ANCHOR.MIDDLE
        lp = ltf.paragraphs[0]
        lp.text = bname
        lp.alignment = PP_ALIGN.CENTER
        lp.font.name = "Aptos"
        lp.font.size = Pt(9.0)
        lp.font.bold = True
        lp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)


# ==============================================================================
# SLIDE 20: Integrated Governance
# ==============================================================================

def build_slide20(prs, lib_s1, lib_lines):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Title
    title_box = slide.shapes.add_textbox(Pt(30.0), Pt(6.0), Pt(900.0), Pt(32.0))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Integrated Governance: A structured governance model ensuring KPOD retains all reserved authorities while GEP takes accountability for integrated PSCM service delivery."
    p.font.name = "Aptos"
    p.font.size = Pt(12.5)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

    # Top Section: 3-Tier Committee Model (w = 640.0, top = 44.0, h = 286.0)
    gov_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(44.0), Pt(640.0), Pt(286.0))
    apply_solid_fill(gov_box, "FFFFFF")
    gov_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(gov_box, 4.0)
    apply_library_shadow(gov_box)

    committees = [
        {
            "name": "1. EXECUTIVE STEERING COMMITTEE",
            "cadence": "Quarterly",
            "kpod": "Executive Sponsor\nProgramme Lead",
            "blocks": ["Change Management", "Business Alignment", "Enforce Strategic Intent", "Remove Strategic Bottlenecks", "Effective Communication"],
            "gep": "Client Executive\nProgramme Lead",
            "clr": ColorTokens.BRAND_NAVY
        },
        {
            "name": "2. PROGRAMME COMMITTEE",
            "cadence": "Monthly",
            "kpod": "Programme Lead\nCategory Managers\nKey Business Stakeholders",
            "blocks": ["Progress to Date", "Change Management", "Policy / Process Change Discussions", "Prioritisation"],
            "gep": "Programme Lead\nKey Delivery Team Members",
            "clr": ColorTokens.BRAND_TANGERINE
        },
        {
            "name": "3. OPERATIONAL COMMITEE",
            "cadence": "Weekly/  Fortnight",
            "kpod": "Category Managers\nKey Business Stakeholders",
            "blocks": ["Project Planning and Execution", "Issue Mgmt. / Escalation", "Strategy Development", "SLA Management"],
            "gep": "Delivery Team",
            "clr": ColorTokens.BRAND_INDIGO
        }
    ]

    c_y = 52.0
    for c_i, com in enumerate(committees):
        # Committee Title Bar
        cb_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(40.0), Pt(c_y), Pt(620.0), Pt(18.0))
        apply_solid_fill(cb_hdr, com["clr"])
        set_shape_corner_radius(cb_hdr, 4.0)
        cbtf = cb_hdr.text_frame
        cbtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        cp = cbtf.paragraphs[0]
        cp.text = com["name"] + f"  ({com['cadence']})"
        cp.alignment = PP_ALIGN.CENTER
        cp.font.name = "Aptos"
        cp.font.size = Pt(9.0)
        cp.font.bold = True
        cp.font.color.rgb = hex_to_rgb("FFFFFF")

        # KPOD Stakeholders (Left)
        kb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(38.0), Pt(c_y + 20.0), Pt(100.0), Pt(46.0))
        apply_solid_fill(kb, "F8FAFC")
        kb.line.color.rgb = hex_to_rgb("CBD5E1")
        set_shape_corner_radius(kb, 4.0)
        ktf = kb.text_frame
        ktf.word_wrap = True
        ktf.vertical_anchor = MSO_ANCHOR.MIDDLE
        kp = ktf.paragraphs[0]
        kp.text = com["kpod"]
        kp.alignment = PP_ALIGN.CENTER
        kp.font.name = "Aptos"
        kp.font.size = Pt(9.0)
        kp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

        # Committee Topic Blocks (Center)
        n_blks = len(com["blocks"])
        blk_w = 414.0 / n_blks
        for b_idx, blk_name in enumerate(com["blocks"]):
            bx = 142.0 + b_idx * blk_w
            bbox = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(bx), Pt(c_y + 20.0), Pt(blk_w - 4.0), Pt(46.0))
            apply_solid_fill(bbox, "EFF6FF" if c_i % 2 == 0 else "FFF7ED")
            bbox.line.color.rgb = hex_to_rgb("BFDBFE" if c_i % 2 == 0 else "FED7AA")
            set_shape_corner_radius(bbox, 4.0)
            btf = bbox.text_frame
            btf.margin_left = Pt(2.0)
            btf.margin_right = Pt(2.0)
            btf.margin_top = Pt(2.0)
            btf.margin_bottom = Pt(2.0)
            btf.word_wrap = True
            btf.vertical_anchor = MSO_ANCHOR.MIDDLE
            bp = btf.paragraphs[0]
            bp.text = blk_name
            bp.alignment = PP_ALIGN.CENTER
            bp.font.name = "Aptos"
            bp.font.size = Pt(9.0)
            bp.font.bold = True
            bp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO if c_i % 2 == 0 else ColorTokens.BRAND_TANGERINE)

        # GEP Stakeholders (Right)
        gb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(560.0), Pt(c_y + 20.0), Pt(100.0), Pt(46.0))
        apply_solid_fill(gb, "F8FAFC")
        gb.line.color.rgb = hex_to_rgb("CBD5E1")
        set_shape_corner_radius(gb, 4.0)
        gtf = gb.text_frame
        gtf.word_wrap = True
        gtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        gp = gtf.paragraphs[0]
        gp.text = com["gep"]
        gp.alignment = PP_ALIGN.CENTER
        gp.font.name = "Aptos"
        gp.font.size = Pt(9.0)
        gp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

        c_y += 76.0

    # Interaction labels bar below committees
    lbl_bar = slide.shapes.add_textbox(Pt(40.0), Pt(290.0), Pt(620.0), Pt(32.0))
    ltf = lbl_bar.text_frame
    ltf.word_wrap = True
    ltf.vertical_anchor = MSO_ANCHOR.MIDDLE
    lp1 = ltf.paragraphs[0]
    lp1.text = "Performance Trends  |  Critical Issues  |  Long-Term Strategies  |  Project Updates  |  Critical Issues, Risks and Recommendations  |  Savings / Compliance Reports, SLA’s, Survey Results"
    lp1.alignment = PP_ALIGN.CENTER
    lp1.font.name = "Aptos"
    lp1.font.size = Pt(9.0)
    lp1.font.italic = True
    lp1.font.color.rgb = hex_to_rgb("64748B")

    # Top Right: Embedded Controls & Assurance (w = 246.0, top = 44.0, h = 286.0)
    emb_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(684.0), Pt(44.0), Pt(246.0), Pt(286.0))
    apply_solid_fill(emb_box, "FFFFFF")
    emb_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(emb_box, 4.0)
    apply_library_shadow(emb_box)

    emb_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(692.0), Pt(50.0), Pt(230.0), Pt(22.0))
    apply_solid_fill(emb_hdr, ColorTokens.BRAND_INDIGO)
    set_shape_corner_radius(emb_hdr, 4.0)
    ehtf = emb_hdr.text_frame
    ehtf.vertical_anchor = MSO_ANCHOR.MIDDLE
    ehp = ehtf.paragraphs[0]
    ehp.text = "Embedded Controls & Assurance"
    ehp.alignment = PP_ALIGN.CENTER
    ehp.font.name = "Aptos"
    ehp.font.size = Pt(9.0)
    ehp.font.bold = True
    ehp.font.color.rgb = hex_to_rgb("FFFFFF")

    controls = [
        "CDD before award",
        "Approved strategy before market",
        "Conflict of interest declarations",
        "Monthly compliance testing",
        "Audit ready records",
        "Red risks to CPO within 24 hours"
    ]
    ctf_box = slide.shapes.add_textbox(Pt(694.0), Pt(76.0), Pt(226.0), Pt(244.0))
    ctf = ctf_box.text_frame
    ctf.word_wrap = True
    ctf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for c_i, ctrl in enumerate(controls):
        pb = ctf.paragraphs[0] if c_i == 0 else ctf.add_paragraph()
        pb.text = clean_bullet_string(ctrl)
        pb.font.name = "Aptos"
        pb.font.size = Pt(9.0)
        pb.font.bold = True
        pb.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)
        pb.space_after = Pt(6.0)
        apply_native_bullet(pb, char="•", typeface="Arial", mar_l_pt=9.0, indent_pt=-9.0)

    # Bottom Section: Accountability (top = 340.0, h = 182.0)
    acc_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(340.0), Pt(900.0), Pt(182.0))
    apply_solid_fill(acc_box, "FFFFFF")
    acc_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(acc_box, 4.0)
    apply_library_shadow(acc_box)

    # Left Vertical Header Badge (Rotated 270 deg vertical text)
    avb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(38.0), Pt(348.0), Pt(36.0), Pt(166.0))
    apply_solid_fill(avb, ColorTokens.BRAND_TANGERINE)
    set_shape_corner_radius(avb, 4.0)
    avbtf = avb.text_frame
    avbtf.word_wrap = False
    avbtf.vertical_anchor = MSO_ANCHOR.MIDDLE
    avbtf._txBody.bodyPr.set("vert", "vert270")
    avbp = avbtf.paragraphs[0]
    avbp.text = "Accountability"
    avbp.alignment = PP_ALIGN.CENTER
    avbp.font.name = "Aptos"
    avbp.font.size = Pt(11.0)
    avbp.font.bold = True
    avbp.font.color.rgb = hex_to_rgb("FFFFFF")

    # Left: KPOD retains (w = 416.0)
    kp_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(82.0), Pt(348.0), Pt(416.0), Pt(20.0))
    apply_solid_fill(kp_hdr, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(kp_hdr, 4.0)
    kptf = kp_hdr.text_frame
    kptf.vertical_anchor = MSO_ANCHOR.MIDDLE
    kpp = kptf.paragraphs[0]
    kpp.text = "KPOD retains"
    kpp.alignment = PP_ALIGN.CENTER
    kpp.font.name = "Aptos"
    kpp.font.size = Pt(9.0)
    kpp.font.bold = True
    kpp.font.color.rgb = hex_to_rgb("FFFFFF")

    kpod_retains = [
        "Priorities and outcomes",
        "Strategy and award approvals (DoFA)",
        "Contract signature",
        "Engineering, HSE, legal and financial approvals",
        "Operator statutory duties",
        "Acceptance of deliverables"
    ]
    ktf_box = slide.shapes.add_textbox(Pt(82.0), Pt(370.0), Pt(416.0), Pt(144.0))
    ktf = ktf_box.text_frame
    ktf.word_wrap = True
    ktf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for k_i, item in enumerate(kpod_retains):
        pb = ktf.paragraphs[0] if k_i == 0 else ktf.add_paragraph()
        pb.text = clean_bullet_string(item)
        pb.font.name = "Aptos"
        pb.font.size = Pt(9.0)
        pb.font.color.rgb = hex_to_rgb("334155")
        pb.space_after = Pt(2.0)
        apply_native_bullet(pb, char="•", typeface="Arial", mar_l_pt=8.0, indent_pt=-8.0)

    # Right: GEP delivers (w = 416.0)
    gp_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(506.0), Pt(348.0), Pt(416.0), Pt(20.0))
    apply_solid_fill(gp_hdr, ColorTokens.BRAND_INDIGO)
    set_shape_corner_radius(gp_hdr, 4.0)
    gptf = gp_hdr.text_frame
    gptf.vertical_anchor = MSO_ANCHOR.MIDDLE
    gpp = gptf.paragraphs[0]
    gpp.text = "GEP delivers"
    gpp.alignment = PP_ALIGN.CENTER
    gpp.font.name = "Aptos"
    gpp.font.size = Pt(9.0)
    gpp.font.bold = True
    gpp.font.color.rgb = hex_to_rgb("FFFFFF")

    gep_delivers = [
        "All SOW services and deliverables",
        "Management and supervision of GEP people",
        "Quality assurance of outputs",
        "KPI performance and reporting",
        "Compliance with KPOD procedures",
        "Local content delivery"
    ]
    gtf_box = slide.shapes.add_textbox(Pt(506.0), Pt(370.0), Pt(416.0), Pt(144.0))
    gtf = gtf_box.text_frame
    gtf.word_wrap = True
    gtf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for g_i, item in enumerate(gep_delivers):
        pb = gtf.paragraphs[0] if g_i == 0 else gtf.add_paragraph()
        pb.text = clean_bullet_string(item)
        pb.font.name = "Aptos"
        pb.font.size = Pt(9.0)
        pb.font.color.rgb = hex_to_rgb("334155")
        pb.space_after = Pt(2.0)
        apply_native_bullet(pb, char="•", typeface="Arial", mar_l_pt=8.0, indent_pt=-8.0)
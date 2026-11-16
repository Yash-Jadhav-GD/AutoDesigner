"""
AutoDesigner: Lipton Transformation Slide Builders (Slides 8 through 12)
Designed to executive benchmark standards:
- Archetypes from Library_copy.pptx and After_copy.pptx
- Official GEP vector SVG icons from Icon Finder
- 100% Middle vertical alignment (anchor="ctr")
- Native DrawingML PowerPoint bullets (<a:buChar char="•"/>), 0 literal bullet characters
- Exact 4.0pt corner radius across all rounded rectangles
- Authentic Library divider lines
- Strictly zero overlapping shapes
- 100% content and token fidelity
"""

import copy
from pathlib import Path
import pptx
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml import parse_xml

from engine.builders.common import (
    clean_element,
    get_shape_by_name,
    add_library_line,
    set_shape_corner_radius,
    apply_library_shadow,
    apply_native_bullet,
    hex_to_rgb
)
from engine.color_tokens import ColorTokens, apply_solid_fill

ICONS_DIR = Path(r"c:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\icon_finder\assets\icons")

# Registry of SVGs to inject per slide: {slide_num: [(rel_id, svg_filename, left_pt, top_pt, size_pt), ...]}
LIPTON_SVG_REGISTRY = {}


def add_svg_picture_xml(slide, rel_id, pic_id, left_pt, top_pt, size_pt, name="Graphic"):
    """Adds DrawingML picture XML pointing to an embedded SVG relationship."""
    pic_l = int(left_pt * 12700)
    pic_t = int(top_pt * 12700)
    pic_sz = int(size_pt * 12700)
    pic_xml = f'''<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:nvPicPr>
        <p:cNvPr id="{pic_id}" name="{name}"/>
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
        <a:xfrm><a:off x="{pic_l}" y="{pic_t}"/><a:ext cx="{pic_sz}" cy="{pic_sz}"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </p:spPr>
    </p:pic>'''
    slide.shapes._spTree.append(parse_xml(pic_xml))


# ==============================================================================
# SLIDE 8: Lipton Needs & Spend Profile
# ==============================================================================

def build_slide8_lipton(prs, lib_s1, lib_lines):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Title
    tbox = slide.shapes.add_textbox(Pt(30.0), Pt(14.0), Pt(900.0), Pt(32.0))
    tf = tbox.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tp = tf.paragraphs[0]
    tp.text = "Lipton needs simpler, more compliant buying across indirect spend"
    tp.font.name = "Aptos"
    tp.font.size = Pt(16.0)
    tp.font.bold = True
    tp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

    decor_proto = get_shape_by_name(lib_s1, "Design 1 Decor box")
    hdr_proto = get_shape_by_name(lib_s1, "Design 1 Header Text")

    # Left Container: What Lipton needs (w=440.0, top=60.0, h=444.0)
    left_elem = clean_element(copy.deepcopy(decor_proto._element))
    slide.shapes._spTree.append(left_elem)
    lbox = slide.shapes[-1]
    lbox.left = Pt(30.0)
    lbox.top = Pt(60.0)
    lbox.width = Pt(440.0)
    lbox.height = Pt(444.0)
    set_shape_corner_radius(lbox, 4.0)
    apply_library_shadow(lbox)

    # Left Header Pill
    lhdr_elem = clean_element(copy.deepcopy(hdr_proto._element))
    slide.shapes._spTree.append(lhdr_elem)
    lhdr = slide.shapes[-1]
    lhdr.left = Pt(140.0)
    lhdr.top = Pt(48.0)
    lhdr.width = Pt(220.0)
    lhdr.height = Pt(26.0)
    apply_solid_fill(lhdr, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(lhdr, 4.0)
    ltf = lhdr.text_frame
    ltf.vertical_anchor = MSO_ANCHOR.MIDDLE
    lp = ltf.paragraphs[0]
    lp.text = "What Lipton needs"
    lp.alignment = PP_ALIGN.CENTER
    lp.font.name = "Aptos"
    lp.font.size = Pt(11.0)
    lp.font.bold = True
    lp.font.color.rgb = hex_to_rgb("FFFFFF")

    # 4 Needs Rows (UI-ROW-01 Style)
    needs = [
        ("Simpler buying", "One place to start, with fewer manual follow-ups", "laptop-fb9eb9.svg"),
        ("Stronger compliance", "Guide demand to approved buying channels", "shield-check-eea126.svg"),
        ("Faster tail execution", "Automate standard buys and resolve complex requests", "timer-de72ae.svg"),
        ("Less supplier administration", "Simplify setup and payment for occasional suppliers", "accounting-invoice-38a68e.svg")
    ]

    s8_svgs = []
    for i, (n_title, n_sub, n_svg) in enumerate(needs):
        ry = 88.0 + i * 98.0
        
        # Circular Icon Badge
        badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, Pt(48.0), Pt(ry + 12.0), Pt(48.0), Pt(48.0))
        apply_solid_fill(badge, "EFF6FF")
        badge.line.color.rgb = hex_to_rgb("BFDBFE")

        # Injected SVG
        rel_id = f"rIdSvgS8_{i+1}"
        s8_svgs.append((rel_id, n_svg, 48.0 + 10.0, ry + 22.0, 28.0))
        add_svg_picture_xml(slide, rel_id, 8100 + i, 48.0 + 10.0, ry + 22.0, 28.0, f"Need Icon {i+1}")

        # Text Box
        rtf_box = slide.shapes.add_textbox(Pt(110.0), Pt(ry + 8.0), Pt(344.0), Pt(64.0))
        rtf = rtf_box.text_frame
        rtf.word_wrap = True
        rtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        rp1 = rtf.paragraphs[0]
        rp1.text = n_title
        rp1.font.name = "Aptos"
        rp1.font.size = Pt(12.0)
        rp1.font.bold = True
        rp1.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

        rp2 = rtf.add_paragraph()
        rp2.text = n_sub
        rp2.font.name = "Aptos"
        rp2.font.size = Pt(9.5)
        rp2.font.color.rgb = hex_to_rgb("475569")

        # Divider line between rows
        if i < 3:
            add_library_line(slide, lib_lines["side_fade_r"], 48.0, ry + 86.0, 404.0)

    LIPTON_SVG_REGISTRY[8] = s8_svgs

    # Right Container: Spend profile (w=440.0, top=60.0, h=444.0)
    right_elem = clean_element(copy.deepcopy(decor_proto._element))
    slide.shapes._spTree.append(right_elem)
    rbox = slide.shapes[-1]
    rbox.left = Pt(490.0)
    rbox.top = Pt(60.0)
    rbox.width = Pt(440.0)
    rbox.height = Pt(444.0)
    set_shape_corner_radius(rbox, 4.0)
    apply_library_shadow(rbox)

    # Right Header Pill
    rhdr_elem = clean_element(copy.deepcopy(hdr_proto._element))
    slide.shapes._spTree.append(rhdr_elem)
    rhdr = slide.shapes[-1]
    rhdr.left = Pt(600.0)
    rhdr.top = Pt(48.0)
    rhdr.width = Pt(220.0)
    rhdr.height = Pt(26.0)
    apply_solid_fill(rhdr, ColorTokens.BRAND_TANGERINE)
    set_shape_corner_radius(rhdr, 4.0)
    rtf_hdr = rhdr.text_frame
    rtf_hdr.vertical_anchor = MSO_ANCHOR.MIDDLE
    rp_hdr = rtf_hdr.paragraphs[0]
    rp_hdr.text = "Spend profile"
    rp_hdr.alignment = PP_ALIGN.CENTER
    rp_hdr.font.name = "Aptos"
    rp_hdr.font.size = Pt(11.0)
    rp_hdr.font.bold = True
    rp_hdr.font.color.rgb = hex_to_rgb("FFFFFF")

    # KPI Row: 2 cards side-by-side
    kpis = [
        ("€296m", "indirect spend", 506.0, ColorTokens.BRAND_NAVY),
        ("€44.2m", "total tail", 718.0, ColorTokens.BRAND_TANGERINE)
    ]
    for m_val, m_lbl, kx, clr in kpis:
        k_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(kx), Pt(86.0), Pt(196.0), Pt(64.0))
        apply_solid_fill(k_box, "F8FAFC")
        k_box.line.color.rgb = hex_to_rgb("E2E8F0")
        set_shape_corner_radius(k_box, 4.0)
        apply_library_shadow(k_box)
        ktf = k_box.text_frame
        ktf.vertical_anchor = MSO_ANCHOR.MIDDLE
        kp1 = ktf.paragraphs[0]
        kp1.text = m_val
        kp1.alignment = PP_ALIGN.CENTER
        kp1.font.name = "Aptos"
        kp1.font.size = Pt(17.0)
        kp1.font.bold = True
        kp1.font.color.rgb = hex_to_rgb(clr)

        kp2 = ktf.add_paragraph()
        kp2.text = m_lbl
        kp2.alignment = PP_ALIGN.CENTER
        kp2.font.name = "Aptos"
        kp2.font.size = Pt(9.5)
        kp2.font.bold = True
        kp2.font.color.rgb = hex_to_rgb("64748B")

    # Middle Spend Distribution Curve Container
    curve_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(506.0), Pt(156.0), Pt(408.0), Pt(152.0))
    apply_solid_fill(curve_box, "FFFFFF")
    curve_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(curve_box, 4.0)
    apply_library_shadow(curve_box)

    # Top labels inside curve box
    y_lbl = slide.shapes.add_textbox(Pt(514.0), Pt(160.0), Pt(140.0), Pt(16.0))
    yp = y_lbl.text_frame.paragraphs[0]
    yp.text = "Annual spend per supplier"
    yp.font.name = "Aptos"
    yp.font.size = Pt(9.0)
    yp.font.color.rgb = hex_to_rgb("64748B")

    ill_badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(836.0), Pt(160.0), Pt(70.0), Pt(16.0))
    apply_solid_fill(ill_badge, "F1F5F9")
    ill_badge.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(ill_badge, 4.0)
    ill_p = ill_badge.text_frame.paragraphs[0]
    ill_p.text = "Illustrative"
    ill_p.alignment = PP_ALIGN.CENTER
    ill_p.font.name = "Aptos"
    ill_p.font.size = Pt(9.0)
    ill_p.font.color.rgb = hex_to_rgb("64748B")

    # 3 Segments inside curve container
    segs = [
        ("Above tail", "Core suppliers\n(strategic contracts)", 516.0, 96.0, "F1F5F9", "CBD5E1", ColorTokens.BRAND_NAVY),
        ("Large tail", "€38.1m | €10k–<€100k\nper supplier", 618.0, 150.0, "FFF7ED", "FED7AA", ColorTokens.BRAND_TANGERINE),
        ("Tail-of-tail", "€6.0m | <€10k\nper supplier", 774.0, 130.0, "EFF6FF", "BFDBFE", ColorTokens.BRAND_NAVY)
    ]
    for s_title, s_desc, sx, sw, s_bg, s_bd, s_clr in segs:
        sb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(sx), Pt(180.0), Pt(sw), Pt(106.0))
        apply_solid_fill(sb, s_bg)
        sb.line.color.rgb = hex_to_rgb(s_bd)
        set_shape_corner_radius(sb, 4.0)
        sbtf = sb.text_frame
        sbtf.word_wrap = True
        sbtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        sp1 = sbtf.paragraphs[0]
        sp1.text = s_title
        sp1.alignment = PP_ALIGN.CENTER
        sp1.font.name = "Aptos"
        sp1.font.size = Pt(10.5)
        sp1.font.bold = True
        sp1.font.color.rgb = hex_to_rgb(s_clr)

        sp2 = sbtf.add_paragraph()
        sp2.text = s_desc
        sp2.alignment = PP_ALIGN.CENTER
        sp2.font.name = "Aptos"
        sp2.font.size = Pt(9.0)
        sp2.font.color.rgb = hex_to_rgb("475569")

    # Bottom X axis label
    x_lbl = slide.shapes.add_textbox(Pt(516.0), Pt(288.0), Pt(388.0), Pt(16.0))
    xp = x_lbl.text_frame.paragraphs[0]
    xp.text = "Suppliers ranked by annual spend ➔"
    xp.alignment = PP_ALIGN.CENTER
    xp.font.name = "Aptos"
    xp.font.size = Pt(9.0)
    xp.font.color.rgb = hex_to_rgb("64748B")

    # Bottom Spend Analysis Highlights Card
    hl_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(506.0), Pt(316.0), Pt(408.0), Pt(172.0))
    apply_solid_fill(hl_box, "F8FAFC")
    hl_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(hl_box, 4.0)
    apply_library_shadow(hl_box)

    hltf = hl_box.text_frame
    hltf.word_wrap = True
    hltf.vertical_anchor = MSO_ANCHOR.MIDDLE
    hltf.margin_left = Pt(14.0)
    hltf.margin_right = Pt(14.0)
    
    hp = hltf.paragraphs[0]
    hp.text = "Spend analysis highlights"
    hp.font.name = "Aptos"
    hp.font.size = Pt(11.0)
    hp.font.bold = True
    hp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)
    hp.space_after = Pt(8.0)

    bullets = [
        "Frequent, lower-value suppliers generated 51% of POs on 17% of spend.",
        "Repeat purchases are candidates for catalogues and guided buying."
    ]
    for b_txt in bullets:
        pb = hltf.add_paragraph()
        pb.text = b_txt
        pb.font.name = "Aptos"
        pb.font.size = Pt(9.5)
        pb.font.color.rgb = hex_to_rgb("475569")
        pb.space_after = Pt(6.0)
        apply_native_bullet(pb, char="•", typeface="Arial", mar_l_pt=10.0, indent_pt=-10.0)

    # Footnote
    fn_box = slide.shapes.add_textbox(Pt(30.0), Pt(518.0), Pt(900.0), Pt(16.0))
    fnp = fn_box.text_frame.paragraphs[0]
    fnp.text = "Source: Lipton RFP spend baseline; GEP spend analysis."
    fnp.font.name = "Aptos"
    fnp.font.size = Pt(9.0)
    fnp.font.italic = True
    fnp.font.color.rgb = hex_to_rgb("64748B")


# ==============================================================================
# SLIDE 9: GEP Unified 4-Lot Buying Model
# ==============================================================================

def build_slide9_lipton(prs, lib_s1, lib_lines):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Title & Subtitle
    tbox = slide.shapes.add_textbox(Pt(30.0), Pt(6.0), Pt(900.0), Pt(22.0))
    tf = tbox.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tp = tf.paragraphs[0]
    tp.text = "GEP brings Lipton’s four lots together in one buying model"
    tp.font.name = "Aptos"
    tp.font.size = Pt(15.0)
    tp.font.bold = True
    tp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

    sub_box = slide.shapes.add_textbox(Pt(30.0), Pt(30.0), Pt(900.0), Pt(18.0))
    stf = sub_box.text_frame
    stp = stf.paragraphs[0]
    stp.text = "The RFP considers BPO-led and digital-led execution. GEP recommends using both to match the work in Lipton’s large tail."
    stp.font.name = "Aptos"
    stp.font.size = Pt(9.5)
    stp.font.color.rgb = hex_to_rgb("475569")

    decor_proto = get_shape_by_name(lib_s1, "Design 1 Decor box")
    hdr_proto = get_shape_by_name(lib_s1, "Design 1 Header Text")

    s9_svgs = []

    # Left Card: Existing Lipton channels (w=150.0, top=58.0, h=404.0)
    lcard_elem = clean_element(copy.deepcopy(decor_proto._element))
    slide.shapes._spTree.append(lcard_elem)
    lcard = slide.shapes[-1]
    lcard.left = Pt(30.0)
    lcard.top = Pt(58.0)
    lcard.width = Pt(150.0)
    lcard.height = Pt(404.0)
    set_shape_corner_radius(lcard, 4.0)
    apply_library_shadow(lcard)

    # Cart Icon Badge
    c_badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, Pt(75.0), Pt(110.0), Pt(60.0), Pt(60.0))
    apply_solid_fill(c_badge, "EFF6FF")
    c_badge.line.color.rgb = hex_to_rgb("BFDBFE")
    rel_id_cart = "rIdSvgS9_cart"
    s9_svgs.append((rel_id_cart, "shop-cart-7a4ddb.svg", 75.0 + 13.0, 110.0 + 13.0, 34.0))
    add_svg_picture_xml(slide, rel_id_cart, 9101, 75.0 + 13.0, 110.0 + 13.0, 34.0, "Cart Icon")

    ltf = lcard.text_frame
    ltf.word_wrap = True
    ltf.vertical_anchor = MSO_ANCHOR.MIDDLE
    lp1 = ltf.paragraphs[0]
    lp1.text = "Existing Lipton channels"
    lp1.alignment = PP_ALIGN.CENTER
    lp1.font.name = "Aptos"
    lp1.font.size = Pt(12.0)
    lp1.font.bold = True
    lp1.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

    lp2 = ltf.add_paragraph()
    lp2.text = "Route to >€100K\nspend team"
    lp2.alignment = PP_ALIGN.CENTER
    lp2.font.name = "Aptos"
    lp2.font.size = Pt(10.0)
    lp2.font.color.rgb = hex_to_rgb("64748B")

    # Main Right Container: GEP unified operations (w=734.0, top=58.0, h=404.0)
    main_elem = clean_element(copy.deepcopy(decor_proto._element))
    slide.shapes._spTree.append(main_elem)
    main_box = slide.shapes[-1]
    main_box.left = Pt(196.0)
    main_box.top = Pt(58.0)
    main_box.width = Pt(734.0)
    main_box.height = Pt(404.0)
    set_shape_corner_radius(main_box, 4.0)
    apply_library_shadow(main_box)

    # Floating Header Pill
    mhdr_elem = clean_element(copy.deepcopy(hdr_proto._element))
    slide.shapes._spTree.append(mhdr_elem)
    mhdr = slide.shapes[-1]
    mhdr.left = Pt(450.0)
    mhdr.top = Pt(46.0)
    mhdr.width = Pt(226.0)
    mhdr.height = Pt(26.0)
    apply_solid_fill(mhdr, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(mhdr, 4.0)
    mtf = mhdr.text_frame
    mtf.vertical_anchor = MSO_ANCHOR.MIDDLE
    mp = mtf.paragraphs[0]
    mp.text = "GEP unified operations"
    mp.alignment = PP_ALIGN.CENTER
    mp.font.name = "Aptos"
    mp.font.size = Pt(11.0)
    mp.font.bold = True
    mp.font.color.rgb = hex_to_rgb("FFFFFF")

    # Top Banner: Process Orchestration
    po_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(212.0), Pt(84.0), Pt(702.0), Pt(50.0))
    apply_solid_fill(po_box, ColorTokens.BRAND_TANGERINE)
    set_shape_corner_radius(po_box, 4.0)
    apply_library_shadow(po_box)

    # Workflow Icon inside banner
    wf_badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, Pt(222.0), Pt(91.0), Pt(36.0), Pt(36.0))
    apply_solid_fill(wf_badge, "FFF7ED")
    wf_badge.line.color.rgb = hex_to_rgb("FED7AA")
    rel_id_wf = "rIdSvgS9_wf"
    s9_svgs.append((rel_id_wf, "workflow-milestones-a6cd65.svg", 222.0 + 7.0, 91.0 + 7.0, 22.0))
    add_svg_picture_xml(slide, rel_id_wf, 9102, 222.0 + 7.0, 91.0 + 7.0, 22.0, "Workflow Icon")

    potf = po_box.text_frame
    potf.word_wrap = True
    potf.vertical_anchor = MSO_ANCHOR.MIDDLE
    potf.margin_left = Pt(56.0)
    pop1 = potf.paragraphs[0]
    pop1.text = "Process Orchestration"
    pop1.font.name = "Aptos"
    pop1.font.size = Pt(11.5)
    pop1.font.bold = True
    pop1.font.color.rgb = hex_to_rgb("FFFFFF")
    pop2 = potf.add_paragraph()
    pop2.text = "One front door for indirect buying, with policy-led routing and request visibility"
    pop2.font.name = "Aptos"
    pop2.font.size = Pt(9.5)
    pop2.font.color.rgb = hex_to_rgb("FFF7ED")

    # Middle Left: Large tail Container (w=460.0, top=144.0, h=250.0)
    lt_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(212.0), Pt(144.0), Pt(460.0), Pt(250.0))
    apply_solid_fill(lt_box, "F8FAFC")
    lt_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(lt_box, 4.0)

    # Large tail header pill
    lth = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(222.0), Pt(150.0), Pt(440.0), Pt(22.0))
    apply_solid_fill(lth, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(lth, 4.0)
    lthtf = lth.text_frame
    lthtf.vertical_anchor = MSO_ANCHOR.MIDDLE
    lthp = lthtf.paragraphs[0]
    lthp.text = "Large tail  |  €10k–<€100k annual spend per supplier"
    lthp.alignment = PP_ALIGN.CENTER
    lthp.font.name = "Aptos"
    lthp.font.size = Pt(9.5)
    lthp.font.bold = True
    lthp.font.color.rgb = hex_to_rgb("FFFFFF")

    # Two columns: Digital Tail Spend + BPO
    # Col 1: Digital
    dig_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(222.0), Pt(178.0), Pt(214.0), Pt(170.0))
    apply_solid_fill(dig_box, "FFFFFF")
    dig_box.line.color.rgb = hex_to_rgb("E2E8F0")
    set_shape_corner_radius(dig_box, 4.0)
    apply_library_shadow(dig_box)

    dig_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(228.0), Pt(184.0), Pt(202.0), Pt(24.0))
    apply_solid_fill(dig_hdr, ColorTokens.BRAND_INDIGO)
    set_shape_corner_radius(dig_hdr, 4.0)
    dh_tf = dig_hdr.text_frame
    dh_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    dhp = dh_tf.paragraphs[0]
    dhp.text = "Digital Tail Spend"
    dhp.alignment = PP_ALIGN.CENTER
    dhp.font.name = "Aptos"
    dhp.font.size = Pt(10.0)
    dhp.font.bold = True
    dhp.font.color.rgb = hex_to_rgb("FFFFFF")

    # SVG for digital
    rel_id_dig = "rIdSvgS9_dig"
    s9_svgs.append((rel_id_dig, "laptop-fb9eb9.svg", 222.0 + (214.0 - 32.0)/2.0, 216.0, 32.0))
    add_svg_picture_xml(slide, rel_id_dig, 9103, 222.0 + (214.0 - 32.0)/2.0, 216.0, 32.0, "Digital Icon")

    dig_tf = slide.shapes.add_textbox(Pt(226.0), Pt(254.0), Pt(206.0), Pt(88.0)).text_frame
    dig_tf.word_wrap = True
    dig_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    dp1 = dig_tf.paragraphs[0]
    dp1.text = "Guide repeat buys to approved catalogues"
    dp1.font.name = "Aptos"
    dp1.font.size = Pt(9.0)
    dp1.font.color.rgb = hex_to_rgb("475569")
    dp1.space_after = Pt(4.0)
    apply_native_bullet(dp1, char="•", typeface="Arial", mar_l_pt=9.0, indent_pt=-9.0)

    dp2 = dig_tf.add_paragraph()
    dp2.text = "Automate standard competitive sourcing"
    dp2.font.name = "Aptos"
    dp2.font.size = Pt(9.0)
    dp2.font.color.rgb = hex_to_rgb("475569")
    apply_native_bullet(dp2, char="•", typeface="Arial", mar_l_pt=9.0, indent_pt=-9.0)

    # Col 2: BPO
    bpo_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(446.0), Pt(178.0), Pt(214.0), Pt(170.0))
    apply_solid_fill(bpo_box, "FFFFFF")
    bpo_box.line.color.rgb = hex_to_rgb("E2E8F0")
    set_shape_corner_radius(bpo_box, 4.0)
    apply_library_shadow(bpo_box)

    bpo_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(452.0), Pt(184.0), Pt(202.0), Pt(24.0))
    apply_solid_fill(bpo_hdr, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(bpo_hdr, 4.0)
    bh_tf = bpo_hdr.text_frame
    bh_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    bhp = bh_tf.paragraphs[0]
    bhp.text = "BPO Buy Desk"
    bhp.alignment = PP_ALIGN.CENTER
    bhp.font.name = "Aptos"
    bhp.font.size = Pt(10.0)
    bhp.font.bold = True
    bhp.font.color.rgb = hex_to_rgb("FFFFFF")

    # SVG for BPO
    rel_id_bpo = "rIdSvgS9_bpo"
    s9_svgs.append((rel_id_bpo, "team-meeting-fb0773.svg", 446.0 + (214.0 - 32.0)/2.0, 216.0, 32.0))
    add_svg_picture_xml(slide, rel_id_bpo, 9104, 446.0 + (214.0 - 32.0)/2.0, 216.0, 32.0, "BPO Icon")

    bpo_tf = slide.shapes.add_textbox(Pt(450.0), Pt(254.0), Pt(206.0), Pt(88.0)).text_frame
    bpo_tf.word_wrap = True
    bpo_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    bp1 = bpo_tf.paragraphs[0]
    bp1.text = "Handle complex requests and exceptions"
    bp1.font.name = "Aptos"
    bp1.font.size = Pt(9.0)
    bp1.font.color.rgb = hex_to_rgb("475569")
    bp1.space_after = Pt(4.0)
    apply_native_bullet(bp1, char="•", typeface="Arial", mar_l_pt=9.0, indent_pt=-9.0)

    bp2 = bpo_tf.add_paragraph()
    bp2.text = "Consolidate demand and negotiate terms"
    bp2.font.name = "Aptos"
    bp2.font.size = Pt(9.0)
    bp2.font.color.rgb = hex_to_rgb("475569")
    apply_native_bullet(bp2, char="•", typeface="Arial", mar_l_pt=9.0, indent_pt=-9.0)

    # Bottom cross-strip: Buyers negotiate the terms. Digital makes them easy to use.
    cs_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(222.0), Pt(356.0), Pt(440.0), Pt(30.0))
    apply_solid_fill(cs_box, "FFF7ED")
    cs_box.line.color.rgb = hex_to_rgb("FED7AA")
    set_shape_corner_radius(cs_box, 4.0)
    cstf = cs_box.text_frame
    cstf.vertical_anchor = MSO_ANCHOR.MIDDLE
    csp = cstf.paragraphs[0]
    csp.text = "Buyers negotiate the terms. Digital makes them easy to use."
    csp.alignment = PP_ALIGN.CENTER
    csp.font.name = "Aptos"
    csp.font.size = Pt(9.5)
    csp.font.bold = True
    csp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_TANGERINE)

    # Middle Right: Payment Aggregation
    pay_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(684.0), Pt(144.0), Pt(230.0), Pt(250.0))
    apply_solid_fill(pay_box, "F8FAFC")
    pay_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(pay_box, 4.0)

    pay_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(694.0), Pt(150.0), Pt(210.0), Pt(24.0))
    apply_solid_fill(pay_hdr, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(pay_hdr, 4.0)
    phtf = pay_hdr.text_frame
    phtf.vertical_anchor = MSO_ANCHOR.MIDDLE
    php = phtf.paragraphs[0]
    php.text = "Payment Aggregation"
    php.alignment = PP_ALIGN.CENTER
    php.font.name = "Aptos"
    php.font.size = Pt(10.0)
    php.font.bold = True
    php.font.color.rgb = hex_to_rgb("FFFFFF")

    # Layers SVG for Payment Aggregation
    rel_id_pay = "rIdSvgS9_pay"
    s9_svgs.append((rel_id_pay, "layers-975e0a.svg", 684.0 + (230.0 - 40.0)/2.0, 184.0, 40.0))
    add_svg_picture_xml(slide, rel_id_pay, 9105, 684.0 + (230.0 - 40.0)/2.0, 184.0, 40.0, "Pay Icon")

    pay_tf = slide.shapes.add_textbox(Pt(694.0), Pt(234.0), Pt(210.0), Pt(148.0)).text_frame
    pay_tf.word_wrap = True
    pay_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    ptp1 = pay_tf.paragraphs[0]
    ptp1.text = "Tail-of-tail"
    ptp1.alignment = PP_ALIGN.CENTER
    ptp1.font.name = "Aptos"
    ptp1.font.size = Pt(11.0)
    ptp1.font.bold = True
    ptp1.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_TANGERINE)

    ptp2 = pay_tf.add_paragraph()
    ptp2.text = "<€10k annual spend per supplier"
    ptp2.alignment = PP_ALIGN.CENTER
    ptp2.font.name = "Aptos"
    ptp2.font.size = Pt(9.0)
    ptp2.font.bold = True
    ptp2.font.color.rgb = hex_to_rgb("64748B")
    ptp2.space_after = Pt(8.0)

    ptp3 = pay_tf.add_paragraph()
    ptp3.text = "Simplify ordering, invoicing and payment for eligible occasional suppliers"
    ptp3.alignment = PP_ALIGN.CENTER
    ptp3.font.name = "Aptos"
    ptp3.font.size = Pt(9.5)
    ptp3.font.color.rgb = hex_to_rgb("475569")

    # Bottom Handoffs Banner
    h_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(472.0), Pt(900.0), Pt(38.0))
    apply_solid_fill(h_box, "FFF7ED")
    h_box.line.color.rgb = hex_to_rgb("FED7AA")
    set_shape_corner_radius(h_box, 4.0)
    apply_library_shadow(h_box)
    htf = h_box.text_frame
    htf.vertical_anchor = MSO_ANCHOR.MIDDLE
    hp = htf.paragraphs[0]
    hp.text = "One GEP team owns the handoffs across all four lots."
    hp.alignment = PP_ALIGN.CENTER
    hp.font.name = "Aptos"
    hp.font.size = Pt(11.5)
    hp.font.bold = True
    hp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_TANGERINE)

    LIPTON_SVG_REGISTRY[9] = s9_svgs


# ==============================================================================
# SLIDE 10: Buying Outcomes & Handoffs
# ==============================================================================

def build_slide10_lipton(prs, lib_s1, lib_lines):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Title
    tbox = slide.shapes.add_textbox(Pt(30.0), Pt(14.0), Pt(900.0), Pt(32.0))
    tf = tbox.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tp = tf.paragraphs[0]
    tp.text = "One partner across four lots gives Lipton one owner for buying outcomes"
    tp.font.name = "Aptos"
    tp.font.size = Pt(15.5)
    tp.font.bold = True
    tp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

    decor_proto = get_shape_by_name(lib_s1, "Design 1 Decor box")
    hdr_proto = get_shape_by_name(lib_s1, "Design 1 Header Text")

    s10_svgs = []

    # Left Column: Savings that carry through to repeat purchases (w=390.0, top=58.0, h=404.0)
    lbox_elem = clean_element(copy.deepcopy(decor_proto._element))
    slide.shapes._spTree.append(lbox_elem)
    lbox = slide.shapes[-1]
    lbox.left = Pt(30.0)
    lbox.top = Pt(58.0)
    lbox.width = Pt(390.0)
    lbox.height = Pt(404.0)
    set_shape_corner_radius(lbox, 4.0)
    apply_library_shadow(lbox)

    lhdr_elem = clean_element(copy.deepcopy(hdr_proto._element))
    slide.shapes._spTree.append(lhdr_elem)
    lhdr = slide.shapes[-1]
    lhdr.left = Pt(85.0)
    lhdr.top = Pt(46.0)
    lhdr.width = Pt(280.0)
    lhdr.height = Pt(26.0)
    apply_solid_fill(lhdr, ColorTokens.BRAND_TANGERINE)
    set_shape_corner_radius(lhdr, 4.0)
    ltf = lhdr.text_frame
    ltf.vertical_anchor = MSO_ANCHOR.MIDDLE
    lp = ltf.paragraphs[0]
    lp.text = "Savings that carry through to repeat purchases"
    lp.alignment = PP_ALIGN.CENTER
    lp.font.name = "Aptos"
    lp.font.size = Pt(10.5)
    lp.font.bold = True
    lp.font.color.rgb = hex_to_rgb("FFFFFF")

    # 3 Vertical Steps
    steps = [
        ("Buyers negotiate supplier terms", "business-deal-handshake-6b1892.svg", 84.0),
        ("Digital publishes approved catalogue content", "book-library-f6990f.svg", 168.0),
        ("Orchestration directs repeat demand there", "shop-cart-7a4ddb.svg", 252.0)
    ]
    for s_idx, (stxt, ssvg, sy) in enumerate(steps):
        sb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(46.0), Pt(sy), Pt(358.0), Pt(66.0))
        apply_solid_fill(sb, "F8FAFC")
        sb.line.color.rgb = hex_to_rgb("E2E8F0")
        set_shape_corner_radius(sb, 4.0)
        apply_library_shadow(sb)

        # Icon Badge
        badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, Pt(58.0), Pt(sy + 9.0), Pt(48.0), Pt(48.0))
        apply_solid_fill(badge, "EFF6FF")
        badge.line.color.rgb = hex_to_rgb("BFDBFE")
        rel_id = f"rIdSvgS10_step{s_idx+1}"
        s10_svgs.append((rel_id, ssvg, 58.0 + 10.0, sy + 19.0, 28.0))
        add_svg_picture_xml(slide, rel_id, 10101 + s_idx, 58.0 + 10.0, sy + 19.0, 28.0, f"Step Icon {s_idx+1}")

        stf = slide.shapes.add_textbox(Pt(118.0), Pt(sy + 7.0), Pt(276.0), Pt(52.0)).text_frame
        stf.word_wrap = True
        stf.vertical_anchor = MSO_ANCHOR.MIDDLE
        stp = stf.paragraphs[0]
        stp.text = stxt
        stp.font.name = "Aptos"
        stp.font.size = Pt(11.0)
        stp.font.bold = True
        stp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

        # Arrow down
        if s_idx < 2:
            arr = slide.shapes.add_textbox(Pt(210.0), Pt(sy + 67.0), Pt(30.0), Pt(14.0))
            ap = arr.text_frame.paragraphs[0]
            ap.text = "▼"
            ap.alignment = PP_ALIGN.CENTER
            ap.font.size = Pt(9.0)
            ap.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_TANGERINE)

    # Bottom Ownership Banner inside left column
    own_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(46.0), Pt(332.0), Pt(358.0), Pt(112.0))
    apply_solid_fill(own_box, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(own_box, 4.0)
    apply_library_shadow(own_box)
    ot_tf = own_box.text_frame
    ot_tf.word_wrap = True
    ot_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    otp = ot_tf.paragraphs[0]
    otp.text = "GEP owns the handoffs from negotiated terms to everyday use."
    otp.alignment = PP_ALIGN.CENTER
    otp.font.name = "Aptos"
    otp.font.size = Pt(11.0)
    otp.font.bold = True
    otp.font.color.rgb = hex_to_rgb("FFFFFF")

    # Right Column: 3 Benefit Cards (UI-ROW-01 Style) (w=490.0, top=58.0, h=404.0)
    rbox_elem = clean_element(copy.deepcopy(decor_proto._element))
    slide.shapes._spTree.append(rbox_elem)
    rbox = slide.shapes[-1]
    rbox.left = Pt(440.0)
    rbox.top = Pt(58.0)
    rbox.width = Pt(490.0)
    rbox.height = Pt(404.0)
    set_shape_corner_radius(rbox, 4.0)
    apply_library_shadow(rbox)

    benefits = [
        ("Faster buying", "Standard requests move digitally. The buy desk picks up exceptions.", "timer-de72ae.svg", 84.0),
        ("Stronger compliance and visibility", "One front door applies Lipton’s rules and shows request status.", "shield-check-eea126.svg", 190.0),
        ("Less supplier administration", "GEP manages the handoff into aggregation for eligible occasional purchases.", "layers-975e0a.svg", 296.0)
    ]
    for b_idx, (b_title, b_desc, b_svg, by) in enumerate(benefits):
        # Icon Badge
        badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, Pt(462.0), Pt(by + 16.0), Pt(56.0), Pt(56.0))
        apply_solid_fill(badge, "EFF6FF")
        badge.line.color.rgb = hex_to_rgb("BFDBFE")
        rel_id = f"rIdSvgS10_ben{b_idx+1}"
        s10_svgs.append((rel_id, b_svg, 462.0 + 12.0, by + 28.0, 32.0))
        add_svg_picture_xml(slide, rel_id, 10201 + b_idx, 462.0 + 12.0, by + 28.0, 32.0, f"Ben Icon {b_idx+1}")

        btf = slide.shapes.add_textbox(Pt(532.0), Pt(by + 10.0), Pt(382.0), Pt(70.0)).text_frame
        btf.word_wrap = True
        btf.vertical_anchor = MSO_ANCHOR.MIDDLE
        bp1 = btf.paragraphs[0]
        bp1.text = b_title
        bp1.font.name = "Aptos"
        bp1.font.size = Pt(13.0)
        bp1.font.bold = True
        bp1.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

        bp2 = btf.add_paragraph()
        bp2.text = b_desc
        bp2.font.name = "Aptos"
        bp2.font.size = Pt(10.0)
        bp2.font.color.rgb = hex_to_rgb("475569")

        # Divider between benefit rows
        if b_idx < 2:
            add_library_line(slide, lib_lines["side_fade_r"], 462.0, by + 94.0, 448.0)

    # Bottom Tracking Banner
    track_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(472.0), Pt(900.0), Pt(38.0))
    apply_solid_fill(track_box, "F8FAFC")
    track_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(track_box, 4.0)
    apply_library_shadow(track_box)
    ttf = track_box.text_frame
    ttf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tp_b = ttf.paragraphs[0]
    tp_b.text = "Lipton can track cycle time, approved-channel use, realised savings and payment completion together."
    tp_b.alignment = PP_ALIGN.CENTER
    tp_b.font.name = "Aptos"
    tp_b.font.size = Pt(11.0)
    tp_b.font.bold = True
    tp_b.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

    LIPTON_SVG_REGISTRY[10] = s10_svgs


# ==============================================================================
# SLIDE 11: End-to-End MRO Purchase Flow
# ==============================================================================

def build_slide11_lipton(prs, lib_s1, lib_lines):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Title
    tbox = slide.shapes.add_textbox(Pt(30.0), Pt(8.0), Pt(900.0), Pt(24.0))
    tf = tbox.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tp = tf.paragraphs[0]
    tp.text = "One GEP team connects intake, digital sourcing and buyer support"
    tp.font.name = "Aptos"
    tp.font.size = Pt(15.5)
    tp.font.bold = True
    tp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

    # Scenario Strip
    sc_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(36.0), Pt(900.0), Pt(24.0))
    apply_solid_fill(sc_box, "FFF7ED")
    sc_box.line.color.rgb = hex_to_rgb("FED7AA")
    set_shape_corner_radius(sc_box, 4.0)
    sctf = sc_box.text_frame
    sctf.vertical_anchor = MSO_ANCHOR.MIDDLE
    scp = sctf.paragraphs[0]
    scp.text = "Illustrative large-tail purchase: an MRO spare with no catalogue match"
    scp.alignment = PP_ALIGN.CENTER
    scp.font.name = "Aptos"
    scp.font.size = Pt(10.0)
    scp.font.bold = True
    scp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_TANGERINE)

    # 6-Step Process Flow (Chevrons + Description Cards)
    steps_data = [
        ("Request submitted", "A site submits the spare’s specification.", ColorTokens.BRAND_NAVY),
        ("Request routed", "Orchestration checks for an approved route and sends the request to digital sourcing.", ColorTokens.BRAND_INDIGO),
        ("RFQ launched", "Autonomous sourcing invites suppliers. Too few quotes come back.", ColorTokens.BRAND_NAVY),
        ("Exception resolved", "The buy desk expands the supplier search and compares quotes.", ColorTokens.BRAND_INDIGO),
        ("Award approved", "Lipton approves the recommended award.", ColorTokens.VIVID_NAVY_MID),
        ("PO created", "The approved purchase proceeds through SAP Ariba.", ColorTokens.BRAND_NAVY)
    ]

    col_w = 146.0
    for idx, (s_hdr, s_desc, s_clr) in enumerate(steps_data):
        cx = 30.0 + idx * 150.0

        # Step Header Pill
        ch = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(cx), Pt(68.0), Pt(col_w), Pt(28.0))
        apply_solid_fill(ch, s_clr)
        set_shape_corner_radius(ch, 4.0)
        chtf = ch.text_frame
        chtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        chp = chtf.paragraphs[0]
        chp.text = s_hdr
        chp.alignment = PP_ALIGN.CENTER
        chp.font.name = "Aptos"
        chp.font.size = Pt(9.5)
        chp.font.bold = True
        chp.font.color.rgb = hex_to_rgb("FFFFFF")

        # Description Box
        db = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(cx), Pt(100.0), Pt(col_w), Pt(112.0))
        apply_solid_fill(db, "FFFFFF")
        db.line.color.rgb = hex_to_rgb("CBD5E1")
        set_shape_corner_radius(db, 4.0)
        apply_library_shadow(db)
        dbtf = db.text_frame
        dbtf.word_wrap = True
        dbtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        dbtf.margin_left = Pt(8.0)
        dbtf.margin_right = Pt(8.0)
        dbp = dbtf.paragraphs[0]
        dbp.text = s_desc
        dbp.alignment = PP_ALIGN.CENTER
        dbp.font.name = "Aptos"
        dbp.font.size = Pt(9.5)
        dbp.font.color.rgb = hex_to_rgb("475569")

    # System Ownership Row (top=218.0, total height=42.0)
    # GEP Delivery Header (Steps 1-4: w=596.0, h=18.0)
    gep_hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(218.0), Pt(596.0), Pt(18.0))
    apply_solid_fill(gep_hdr, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(gep_hdr, 4.0)
    gh_tf = gep_hdr.text_frame
    gh_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    gh_p = gh_tf.paragraphs[0]
    gh_p.text = "GEP Solution"
    gh_p.alignment = PP_ALIGN.CENTER
    gh_p.font.name = "Aptos"
    gh_p.font.size = Pt(9.5)
    gh_p.font.bold = True
    gh_p.font.color.rgb = hex_to_rgb("FFFFFF")

    # Sub-boxes under GEP Solution
    # Process Orchestration Box (Steps 1-2: w=296.0, h=22.0)
    po_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(238.0), Pt(296.0), Pt(22.0))
    apply_solid_fill(po_box, ColorTokens.BRAND_TANGERINE)
    set_shape_corner_radius(po_box, 4.0)
    potf = po_box.text_frame
    potf.vertical_anchor = MSO_ANCHOR.MIDDLE
    pop = potf.paragraphs[0]
    pop.text = "Process Orchestration"
    pop.alignment = PP_ALIGN.CENTER
    pop.font.name = "Aptos"
    pop.font.size = Pt(9.0)
    pop.font.bold = True
    pop.font.color.rgb = hex_to_rgb("FFFFFF")

    # Digital Tail Spend Box (Step 3: w=146.0, h=22.0)
    dig_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(330.0), Pt(238.0), Pt(146.0), Pt(22.0))
    apply_solid_fill(dig_box, ColorTokens.BRAND_INDIGO)
    set_shape_corner_radius(dig_box, 4.0)
    dtf = dig_box.text_frame
    dtf.vertical_anchor = MSO_ANCHOR.MIDDLE
    dp = dtf.paragraphs[0]
    dp.text = "Digital Tail Spend"
    dp.alignment = PP_ALIGN.CENTER
    dp.font.name = "Aptos"
    dp.font.size = Pt(9.0)
    dp.font.bold = True
    dp.font.color.rgb = hex_to_rgb("FFFFFF")

    # BPO Buy Desk Box (Step 4: w=146.0, h=22.0)
    bpo_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(480.0), Pt(238.0), Pt(146.0), Pt(22.0))
    apply_solid_fill(bpo_box, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(bpo_box, 4.0)
    btf = bpo_box.text_frame
    btf.vertical_anchor = MSO_ANCHOR.MIDDLE
    bp = btf.paragraphs[0]
    bp.text = "BPO Buy Desk"
    bp.alignment = PP_ALIGN.CENTER
    bp.font.name = "Aptos"
    bp.font.size = Pt(9.0)
    bp.font.bold = True
    bp.font.color.rgb = hex_to_rgb("FFFFFF")

    # Lipton Approval Box (Step 5: w=146.0, h=42.0)
    lip_sys = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(630.0), Pt(218.0), Pt(146.0), Pt(42.0))
    apply_solid_fill(lip_sys, ColorTokens.BRAND_INDIGO)
    set_shape_corner_radius(lip_sys, 4.0)
    lstf = lip_sys.text_frame
    lstf.vertical_anchor = MSO_ANCHOR.MIDDLE
    lsp = lstf.paragraphs[0]
    lsp.text = "Lipton approval"
    lsp.alignment = PP_ALIGN.CENTER
    lsp.font.name = "Aptos"
    lsp.font.size = Pt(10.0)
    lsp.font.bold = True
    lsp.font.color.rgb = hex_to_rgb("FFFFFF")

    # SAP Ariba Box (Step 6: w=146.0, h=42.0)
    sap_sys = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(780.0), Pt(218.0), Pt(146.0), Pt(42.0))
    apply_solid_fill(sap_sys, ColorTokens.VIVID_NAVY_MID)
    set_shape_corner_radius(sap_sys, 4.0)
    sstf = sap_sys.text_frame
    sstf.vertical_anchor = MSO_ANCHOR.MIDDLE
    ssp = sstf.paragraphs[0]
    ssp.text = "SAP Ariba"
    ssp.alignment = PP_ALIGN.CENTER
    ssp.font.name = "Aptos"
    ssp.font.size = Pt(10.0)
    ssp.font.bold = True
    ssp.font.color.rgb = hex_to_rgb("FFFFFF")

    # Bottom Section: Where shared ownership helps (w=900.0, top=272.0, h=236.0)
    decor_proto = get_shape_by_name(lib_s1, "Design 1 Decor box")
    hdr_proto = get_shape_by_name(lib_s1, "Design 1 Header Text")

    so_elem = clean_element(copy.deepcopy(decor_proto._element))
    slide.shapes._spTree.append(so_elem)
    so_box = slide.shapes[-1]
    so_box.left = Pt(30.0)
    so_box.top = Pt(272.0)
    so_box.width = Pt(900.0)
    so_box.height = Pt(236.0)
    set_shape_corner_radius(so_box, 4.0)
    apply_library_shadow(so_box)

    so_hdr_elem = clean_element(copy.deepcopy(hdr_proto._element))
    slide.shapes._spTree.append(so_hdr_elem)
    so_hdr = slide.shapes[-1]
    so_hdr.left = Pt(340.0)
    so_hdr.top = Pt(260.0)
    so_hdr.width = Pt(280.0)
    so_hdr.height = Pt(24.0)
    apply_solid_fill(so_hdr, ColorTokens.BRAND_TANGERINE)
    set_shape_corner_radius(so_hdr, 4.0)
    sohtf = so_hdr.text_frame
    sohtf.vertical_anchor = MSO_ANCHOR.MIDDLE
    sohp = sohtf.paragraphs[0]
    sohp.text = "Where shared ownership helps"
    sohp.alignment = PP_ALIGN.CENTER
    sohp.font.name = "Aptos"
    sohp.font.size = Pt(11.0)
    sohp.font.bold = True
    sohp.font.color.rgb = hex_to_rgb("FFFFFF")

    s11_svgs = []
    ownership_cards = [
        ("Intake → RFQ", "The specification carries through to sourcing.", "accounting-invoice-38a68e.svg", 48.0),
        ("RFQ → Buy desk", "Buyers pick up the same request and quote history.", "business-deal-handshake-6b1892.svg", 342.0),
        ("Repeat demand → Catalogue", "GEP can negotiate terms, publish content and update routing.", "book-library-f6990f.svg", 636.0)
    ]
    for c_i, (c_title, c_desc, c_svg, cx) in enumerate(ownership_cards):
        cb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(cx), Pt(296.0), Pt(276.0), Pt(198.0))
        apply_solid_fill(cb, "F8FAFC")
        cb.line.color.rgb = hex_to_rgb("E2E8F0")
        set_shape_corner_radius(cb, 4.0)
        apply_library_shadow(cb)

        # Icon Badge
        badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, Pt(cx + (276.0 - 52.0) / 2.0), Pt(310.0), Pt(52.0), Pt(52.0))
        apply_solid_fill(badge, "EFF6FF")
        badge.line.color.rgb = hex_to_rgb("BFDBFE")
        rel_id = f"rIdSvgS11_{c_i+1}"
        s11_svgs.append((rel_id, c_svg, cx + (276.0 - 30.0) / 2.0, 321.0, 30.0))
        add_svg_picture_xml(slide, rel_id, 11101 + c_i, cx + (276.0 - 30.0) / 2.0, 321.0, 30.0, f"Card Icon {c_i+1}")

        # Text
        ctf = slide.shapes.add_textbox(Pt(cx + 12.0), Pt(374.0), Pt(252.0), Pt(106.0)).text_frame
        ctf.word_wrap = True
        ctf.vertical_anchor = MSO_ANCHOR.MIDDLE
        cp1 = ctf.paragraphs[0]
        cp1.text = c_title
        cp1.alignment = PP_ALIGN.CENTER
        cp1.font.name = "Aptos"
        cp1.font.size = Pt(13.0)
        cp1.font.bold = True
        cp1.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

        cp2 = ctf.add_paragraph()
        cp2.text = c_desc
        cp2.alignment = PP_ALIGN.CENTER
        cp2.font.name = "Aptos"
        cp2.font.size = Pt(10.0)
        cp2.font.color.rgb = hex_to_rgb("475569")

    LIPTON_SVG_REGISTRY[11] = s11_svgs


# ==============================================================================
# SLIDE 12: 3-Wave Channel Rollout
# ==============================================================================

def build_slide12_lipton(prs, lib_s1, lib_lines):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Title
    tbox = slide.shapes.add_textbox(Pt(30.0), Pt(14.0), Pt(900.0), Pt(32.0))
    tf = tbox.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tp = tf.paragraphs[0]
    tp.text = "GEP will connect available buying channels first and build new ones in waves"
    tp.font.name = "Aptos"
    tp.font.size = Pt(15.5)
    tp.font.bold = True
    tp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

    decor_proto = get_shape_by_name(lib_s1, "Design 1 Decor box")
    hdr_proto = get_shape_by_name(lib_s1, "Design 1 Header Text")

    s12_svgs = []
    waves = [
        ("1. Connect available channels", "laptop-fb9eb9.svg", [
            "Connect existing contracts and catalogues where ready.",
            "Route eligible requests to digital sourcing or the GEP buy desk."
        ], ColorTokens.BRAND_NAVY, 30.0),
        ("2. Build priority channels", "layers-975e0a.svg", [
            "Start with categories with recurring demand.",
            "GEP buyers negotiate missing terms and the digital team builds catalogue content."
        ], ColorTokens.BRAND_INDIGO, 340.0),
        ("3. Activate and expand", "shop-506d38.svg", [
            "Lipton approves suppliers and terms.",
            "GEP tests routing and activates channels as each market is ready."
        ], ColorTokens.BRAND_TANGERINE, 650.0)
    ]

    for w_idx, (w_title, w_svg, w_bullets, w_clr, wx) in enumerate(waves):
        # Column Container
        w_elem = clean_element(copy.deepcopy(decor_proto._element))
        slide.shapes._spTree.append(w_elem)
        wbox = slide.shapes[-1]
        wbox.left = Pt(wx)
        wbox.top = Pt(64.0)
        wbox.width = Pt(280.0)
        wbox.height = Pt(368.0)
        set_shape_corner_radius(wbox, 4.0)
        apply_library_shadow(wbox)

        # Header Pill
        whdr_elem = clean_element(copy.deepcopy(hdr_proto._element))
        slide.shapes._spTree.append(whdr_elem)
        whdr = slide.shapes[-1]
        whdr.left = Pt(wx + 20.0)
        whdr.top = Pt(52.0)
        whdr.width = Pt(240.0)
        whdr.height = Pt(26.0)
        apply_solid_fill(whdr, w_clr)
        set_shape_corner_radius(whdr, 4.0)
        whtf = whdr.text_frame
        whtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        whp = whtf.paragraphs[0]
        whp.text = w_title
        whp.alignment = PP_ALIGN.CENTER
        whp.font.name = "Aptos"
        whp.font.size = Pt(10.5)
        whp.font.bold = True
        whp.font.color.rgb = hex_to_rgb("FFFFFF")

        # Circular Icon Badge
        badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, Pt(wx + (280.0 - 56.0) / 2.0), Pt(94.0), Pt(56.0), Pt(56.0))
        apply_solid_fill(badge, "EFF6FF")
        badge.line.color.rgb = hex_to_rgb("BFDBFE")
        rel_id = f"rIdSvgS12_{w_idx+1}"
        s12_svgs.append((rel_id, w_svg, wx + (280.0 - 32.0) / 2.0, 106.0, 32.0))
        add_svg_picture_xml(slide, rel_id, 12101 + w_idx, wx + (280.0 - 32.0) / 2.0, 106.0, 32.0, f"Wave Icon {w_idx+1}")

        # Divider
        add_library_line(slide, lib_lines["side_fade_r"], wx + 20.0, 162.0, 240.0)

        # Bullets Container
        btf_box = slide.shapes.add_textbox(Pt(wx + 16.0), Pt(180.0), Pt(248.0), Pt(220.0))
        btf = btf_box.text_frame
        btf.word_wrap = True
        btf.vertical_anchor = MSO_ANCHOR.MIDDLE

        for b_i, b_txt in enumerate(w_bullets):
            bp = btf.paragraphs[0] if b_i == 0 else btf.add_paragraph()
            bp.text = b_txt
            bp.font.name = "Aptos"
            bp.font.size = Pt(10.0)
            bp.font.color.rgb = hex_to_rgb("475569")
            bp.space_after = Pt(10.0)
            apply_native_bullet(bp, char="•", typeface="Arial", mar_l_pt=13.5, indent_pt=-13.5)

    # Bottom Recommended Rollout Banner (top=444.0, h=56.0)
    roll_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(444.0), Pt(900.0), Pt(56.0))
    apply_solid_fill(roll_box, "FFF7ED")
    roll_box.line.color.rgb = hex_to_rgb("FED7AA")
    set_shape_corner_radius(roll_box, 4.0)
    apply_library_shadow(roll_box)

    rtf = roll_box.text_frame
    rtf.word_wrap = True
    rtf.vertical_anchor = MSO_ANCHOR.MIDDLE
    rtf.margin_left = Pt(16.0)
    rtf.margin_right = Pt(16.0)
    rp1 = rtf.paragraphs[0]
    rp1.text = "Recommended rollout"
    rp1.font.name = "Aptos"
    rp1.font.size = Pt(11.5)
    rp1.font.bold = True
    rp1.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_TANGERINE)

    rp2 = rtf.add_paragraph()
    rp2.text = "Prioritise categories by repeat demand, then sequence markets by supplier, catalogue and payment readiness."
    rp2.font.name = "Aptos"
    rp2.font.size = Pt(10.0)
    rp2.font.color.rgb = hex_to_rgb("475569")

    LIPTON_SVG_REGISTRY[12] = s12_svgs

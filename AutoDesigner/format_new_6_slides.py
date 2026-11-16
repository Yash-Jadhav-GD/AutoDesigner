"""
AutoDesigner Precision Formatting Engine for All 6 Slides in Design tester new.pptx
Strictly adheres to:
1. Lines from Library_copy Slide 6:
   - Grey Side Fade R (horizontal fade right)
   - Grey Centre Fade H (horizontal center fade)
   - Grey Solid H (horizontal solid)
   - Grey Solid V (vertical solid)
2. COLOR_THEORY_SPECIFICATION.md & After_copy.pptx visual benchmarks:
   - Slide 1: 3-Part transformation narrative (Problem -> Bridge -> Future)
   - Slide 2: Unified vivid.indigo.soft (#B1A8D8) backplates for parallel diagnostic points + Tangerine INSIGHTS callout + Embedded SVG vector icons
   - Slide 3: Full-width unified vivid.indigo.soft (#B1A8D8) backplates
   - Slide 4: Case Study 3: Opportunity Assessment (Top Scope Bar + 3 Columns with Tangerine Key Opportunities + Bottom Key Outcomes Bar)
   - Slide 5: QXO Leadership Priorities: 8x4 DrawingML Data Table with Indigo headers, zebra striping, and cell borders
   - Slide 6: Opportunity Perspective: 4 Workstreams (A, B, C, D) with 5-phase chromatic ramp, 4-column structured grid, and impact metrics
3. 100% Text Fidelity (ZERO added words, ZERO subtracted words)
4. Exactly 6 slides in output matching 6 input slides
"""

import copy
import io
import os
import shutil
import sys
import zipfile
from xml.etree import ElementTree as ET
from pathlib import Path

import pptx
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.oxml import parse_xml

BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))

from engine.builders.common import (
    clean_element,
    get_shape_by_name,
    add_library_line,
    load_library,
    set_cell_border,
    add_connector_line,
    inject_svg_assets,
    set_shape_corner_radius,
    enforce_deck_corner_radii,
    apply_library_shadow,
)
from engine.color_tokens import (
    ColorTokens,
    hex_to_rgb,
    apply_linear_gradient,
    apply_solid_fill,
)

# Import legacy builders 1, 2, 3
from scratch.legacy.format_all_6_slides import (
    extract_slide1_data,
    build_slide1,
    extract_slide2_data,
    build_slide2,
    extract_slide3_data,
    build_slide3,
)

INPUT_PATH = BASE_DIR / "Design_tester_new_inspect.pptx"
OUTPUT_PATH = BASE_DIR / "output" / "Design_tester_new_formatted.pptx"
LOCAL_COPY_PATH = BASE_DIR / "Design tester new_formatted.pptx"


def extract_slide4_data(input_prs):
    s = input_prs.slides[3]
    title = [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Case Study 3" in p.text][0]
    top_header = [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Accelerated Sourcing Program" in p.text][0]
    
    scope_items = [
        [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Spend Assessed" in p.text][0],
        [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Categories:" in p.text][0],
        [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Regions-in-scope" in p.text][0],
        [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Opportunities Identified" in p.text][0],
    ]
    
    pain_header = [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Pain Points" in p.text][0]
    pain_bullets = [
        "Decentralized sourcing practices across OpCos limited enterprise-wide leverage",
        "Fragmented spend and supplier information constrained enterprise visibility",
        "Predominantly reactive sourcing with limited forward-looking category pipelines",
        "Inconsistent supplier arrangements and sourcing practices across the organization",
        "Limited use of category intelligence and market insights in sourcing decisions",
        "Recent acquisitions created targeted contract harmonization opportunities",
    ]
    
    approach_header = [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if p.text.strip() == "Program Approach"][0]
    app_sec1_title = [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Assess and Prioritize" in p.text][0]
    app_sec1_bullets = [
        "Analyzed spend, supplier arrangements and category requirements",
        "Validated opportunities with procurement and OpCo stakeholders",
        "Prioritized initiatives based on value, feasibility and business readiness",
    ]
    app_sec2_title = [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Develop and Execute" in p.text][0]
    app_sec2_bullets = [
        "Developed category-specific sourcing and negotiation strategies",
        "Mobilized priority initiatives with business stakeholders",
        "Tracked savings through negotiation and implementation",
    ]
    
    opps_header = [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Key Opportunities" in p.text][0]
    opps_items = [
        "IT contract harmonization: Reviewed supplier contracts and identified opportunities to move acquired businesses to more advantageous enterprise agreements",
        "Acquisition synergies: Compared commercial arrangements across the firm and identified pricing and consolidation opportunities post acquisitions",
        "Fleet transformation: Developed a broader operating-model and transformation perspective for global fleet leveraging full scale of combined companies",
        "Strategic sourcing pipeline: Established a prioritized portfolio of category initiatives aligned with business requirements identified through synergy opportunity analysis as well as broader opportunity assessment",
        "Identified and partnered with critical stakeholders to enable successful execution against identified opportunities",
    ]
    
    outcomes_header = [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Key Outcomes" in p.text][0]
    outcomes_metrics = [
        "5%-6% annual uplift in savings, 14% - 15% savings via multi-year strategies",
        "Shift from Reactive to Proactive Sourcing",
        "Prioritized Sourcing Pipeline Developed",
        "~700 Supplier Contracts Renewed",
    ]
    
    return {
        "title": title,
        "top_header": top_header,
        "scope_items": scope_items,
        "pain_header": pain_header,
        "pain_bullets": pain_bullets,
        "approach_header": approach_header,
        "app_sec1_title": app_sec1_title,
        "app_sec1_bullets": app_sec1_bullets,
        "app_sec2_title": app_sec2_title,
        "app_sec2_bullets": app_sec2_bullets,
        "opps_header": opps_header,
        "opps_items": opps_items,
        "outcomes_header": outcomes_header,
        "outcomes_metrics": outcomes_metrics,
    }


def extract_slide5_data(input_prs):
    s = input_prs.slides[4]
    title = [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "QXO" in p.text][0]
    
    headers = [
        "Category",
        "Spend Managed by GEP",
        "Subject Matter Experts",
        "Transformative Value Levers (more detail in appendix)",
    ]
    
    table_shape = [sh for sh in s.shapes if sh.has_table][0]
    t = table_shape.table
    rows = []
    for r in t.rows:
        row_vals = []
        for c in r.cells:
            cell_paras = [p.text.strip() for p in c.text_frame.paragraphs if p.text.strip()]
            row_vals.append(cell_paras)
        rows.append(row_vals)
        
    return {
        "title": title,
        "headers": headers,
        "rows": rows,
    }


def extract_slide6_data(input_prs):
    s = input_prs.slides[5]
    title = [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Opportunity Perspective" in p.text][0]
    headers = [
        [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if p.text.strip() == "Workstreams"][0],
        [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Spend" in p.text and "Suppliers" in p.text][0],
        [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if p.text.strip() == "Planned Interventions"][0],
        [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if p.text.strip() == "Estimated Impact"][0],
    ]
    
    tracks = [
        {
            "letter": "A",
            "name": [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Cross-" in p.text][0],
            "spend_sup": [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Overlapping" in p.text or "$166M" in p.text],
            "interventions": [
                "Consolidate spend & commercial terms across acquired businesses",
                "Renegotiate enterprise-wide agreements leveraging combined scale",
                "Harmonize pricing, rebates, payment terms, and service levels",
            ],
            "impact": [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "$20" in p.text or "(12-15%)" in p.text][:2],
        },
        {
            "letter": "B",
            "name": [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Strategic Sourcing" in p.text][0],
            "spend_sup": [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "> $5M" in p.text or "$450M" in p.text],
            "interventions": [
                "Reassess category strategies using  combined QXO spend footprint",
                "Run targeted sourcing waves for high-value categories and suppliers",
                "Introduce competitive tension and marketplace benchmarking",
            ],
            "impact": [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "$55" in p.text or ("(12-15%)" in p.text and sh.top > 200 * 12700)][:2],
        },
        {
            "letter": "C",
            "name": [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Accelerated RFX" in p.text][0],
            "spend_sup": [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "$500k - $5M" in p.text or "$305M" in p.text],
            "interventions": [
                "Stand up a centralized RFx execution engine across New QXO",
                "Execute sourcing waves at scale using AI-enabled sourcing infrastructure",
                "Provide dedicated sourcing capacity to address opportunities that would otherwise remain untouched",
            ],
            "impact": [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "$15" in p.text or "(5-6%)" in p.text][:2],
        },
        {
            "letter": "D",
            "name": [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Digital Tail" in p.text][0],
            "spend_sup": [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "< $500K" in p.text or "$310M" in p.text],
            "interventions": [
                "Create a centralized procurement channel for  spend across New QXO",
                "Free up QXO Indirect leader time to focus on integration priorities",
                "Enable scalable execution through a combination of managed services, digital workflows and AI",
            ],
            "impact": [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "$10-15M" in p.text or "(3-4%)" in p.text][:2],
        },
    ]
    
    return {
        "title": title,
        "headers": headers,
        "tracks": tracks,
    }


def build_slide4(prs, d, lib_lines):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # 1. Title
    title_box = slide.shapes.add_textbox(Pt(30.0), Pt(10.0), Pt(900.0), Pt(36.0))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = d["title"]
    p.font.name = "Aptos"
    p.font.size = Pt(14.5)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

    # 2. Top Program Scope Box
    scope_box = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(48.0), Pt(900.0), Pt(58.0))
    apply_solid_fill(scope_box, "F8FAFC")
    scope_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(scope_box, 4.0)
    apply_library_shadow(scope_box)

    # Header inside scope box
    hdr_box = slide.shapes.add_textbox(Pt(30.0), Pt(50.0), Pt(900.0), Pt(20.0))
    h_p = hdr_box.text_frame.paragraphs[0]
    h_p.text = d["top_header"]
    h_p.alignment = pptx.enum.text.PP_ALIGN.CENTER
    h_p.font.name = "Aptos"
    h_p.font.size = Pt(10.5)
    h_p.font.bold = True
    h_p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)

    add_library_line(slide, lib_lines["side_fade_r"], 45.0, 70.0, 870.0)

    # 4 Scope tiles
    scope_widths = [210.0, 240.0, 190.0, 220.0]
    scope_lefts = [45.0, 260.0, 505.0, 700.0]
    for i, s_txt in enumerate(d["scope_items"]):
        s_box = slide.shapes.add_textbox(Pt(scope_lefts[i]), Pt(72.0), Pt(scope_widths[i]), Pt(28.0))
        s_p = s_box.text_frame.paragraphs[0]
        s_p.text = s_txt
        s_p.alignment = pptx.enum.text.PP_ALIGN.CENTER
        s_p.font.name = "Aptos"
        s_p.font.size = Pt(9.0)
        s_p.font.bold = True
        s_p.font.color.rgb = hex_to_rgb("334155")

    # 3. Middle 3 Columns
    col_w = 290.0
    col_h = 312.0
    col_top = 112.0
    col_lefts = [30.0, 335.0, 640.0]

    # Col 1: Pain Points / Challenges
    c1_card = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(col_lefts[0]), Pt(col_top), Pt(col_w), Pt(col_h))
    apply_solid_fill(c1_card, "FFFFFF")
    c1_card.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(c1_card, 4.0)
    apply_library_shadow(c1_card)

    c1_tab = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(col_lefts[0]), Pt(col_top), Pt(col_w), Pt(26.0))
    apply_solid_fill(c1_tab, ColorTokens.BRAND_INDIGO)
    c1_tab.line.fill.background()
    set_shape_corner_radius(c1_tab, 4.0)
    t1_p = c1_tab.text_frame.paragraphs[0]
    t1_p.text = d["pain_header"]
    t1_p.alignment = pptx.enum.text.PP_ALIGN.CENTER
    t1_p.font.name = "Aptos"
    t1_p.font.size = Pt(10.5)
    t1_p.font.bold = True
    t1_p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    add_library_line(slide, lib_lines["side_fade_r"], col_lefts[0] + 10.0, col_top + 26.0, 270.0)

    # 6 Pain points
    p_base_y = col_top + 32.0
    p_step_y = 45.0
    for p_i, bullet_text in enumerate(d["pain_bullets"]):
        b_box = slide.shapes.add_textbox(Pt(col_lefts[0] + 10.0), Pt(p_base_y + p_i * p_step_y), Pt(col_w - 20.0), Pt(42.0))
        b_tf = b_box.text_frame
        b_tf.word_wrap = True
        b_tf.margin_top = Pt(1.0)
        b_tf.margin_bottom = Pt(1.0)
        b_p = b_tf.paragraphs[0]
        b_p.text = "• " + bullet_text
        b_p.font.name = "Aptos"
        b_p.font.size = Pt(9.5)
        b_p.font.color.rgb = hex_to_rgb("334155")
        if p_i < 5:
            add_library_line(slide, lib_lines["side_fade_r"], col_lefts[0] + 15.0, p_base_y + (p_i + 1) * p_step_y - 2.0, col_w - 30.0)

    # Col 2: Program Approach
    c2_card = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(col_lefts[1]), Pt(col_top), Pt(col_w), Pt(col_h))
    apply_solid_fill(c2_card, "FFFFFF")
    c2_card.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(c2_card, 4.0)
    apply_library_shadow(c2_card)

    c2_tab = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(col_lefts[1]), Pt(col_top), Pt(col_w), Pt(26.0))
    apply_solid_fill(c2_tab, ColorTokens.BRAND_INDIGO)
    c2_tab.line.fill.background()
    set_shape_corner_radius(c2_tab, 4.0)
    t2_p = c2_tab.text_frame.paragraphs[0]
    t2_p.text = d["approach_header"]
    t2_p.alignment = pptx.enum.text.PP_ALIGN.CENTER
    t2_p.font.name = "Aptos"
    t2_p.font.size = Pt(10.5)
    t2_p.font.bold = True
    t2_p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    add_library_line(slide, lib_lines["side_fade_r"], col_lefts[1] + 10.0, col_top + 26.0, 270.0)

    # Sub-sec 1
    sec1_box = slide.shapes.add_textbox(Pt(col_lefts[1] + 10.0), Pt(col_top + 32.0), Pt(col_w - 20.0), Pt(20.0))
    s1_p = sec1_box.text_frame.paragraphs[0]
    s1_p.text = d["app_sec1_title"]
    s1_p.font.name = "Aptos"
    s1_p.font.size = Pt(9.5)
    s1_p.font.bold = True
    s1_p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)

    add_library_line(slide, lib_lines["side_fade_r"], col_lefts[1] + 10.0, col_top + 52.0, 270.0)

    b_box1 = slide.shapes.add_textbox(Pt(col_lefts[1] + 10.0), Pt(col_top + 54.0), Pt(col_w - 20.0), Pt(80.0))
    b_tf1 = b_box1.text_frame
    b_tf1.word_wrap = True
    for idx, b_txt in enumerate(d["app_sec1_bullets"]):
        p = b_tf1.paragraphs[0] if idx == 0 else b_tf1.add_paragraph()
        p.text = "• " + b_txt
        p.font.name = "Aptos"
        p.font.size = Pt(9.0)
        p.font.color.rgb = hex_to_rgb("334155")
        p.space_after = Pt(2.0)

    # Sub-sec 2
    sec2_top = col_top + 168.0
    sec2_box = slide.shapes.add_textbox(Pt(col_lefts[1] + 10.0), Pt(sec2_top), Pt(col_w - 20.0), Pt(20.0))
    s2_p = sec2_box.text_frame.paragraphs[0]
    s2_p.text = d["app_sec2_title"]
    s2_p.font.name = "Aptos"
    s2_p.font.size = Pt(9.5)
    s2_p.font.bold = True
    s2_p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)

    add_library_line(slide, lib_lines["side_fade_r"], col_lefts[1] + 10.0, sec2_top + 20.0, 270.0)

    b_box2 = slide.shapes.add_textbox(Pt(col_lefts[1] + 10.0), Pt(sec2_top + 22.0), Pt(col_w - 20.0), Pt(80.0))
    b_tf2 = b_box2.text_frame
    b_tf2.word_wrap = True
    for idx, b_txt in enumerate(d["app_sec2_bullets"]):
        p = b_tf2.paragraphs[0] if idx == 0 else b_tf2.add_paragraph()
        p.text = "• " + b_txt
        p.font.name = "Aptos"
        p.font.size = Pt(9.0)
        p.font.color.rgb = hex_to_rgb("334155")
        p.space_after = Pt(2.0)

    # Col 3: Key Opportunities (Tangerine focal callout)
    c3_card = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(col_lefts[2]), Pt(col_top), Pt(col_w), Pt(col_h))
    apply_solid_fill(c3_card, "FFFFFF")
    c3_card.line.color.rgb = hex_to_rgb(ColorTokens.BRAND_TANGERINE)
    c3_card.line.width = Pt(1.5)
    set_shape_corner_radius(c3_card, 4.0)
    apply_library_shadow(c3_card)

    c3_tab = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(col_lefts[2]), Pt(col_top), Pt(col_w), Pt(26.0))
    apply_solid_fill(c3_tab, ColorTokens.BRAND_TANGERINE)
    c3_tab.line.fill.background()
    set_shape_corner_radius(c3_tab, 4.0)
    t3_p = c3_tab.text_frame.paragraphs[0]
    t3_p.text = d["opps_header"]
    t3_p.alignment = pptx.enum.text.PP_ALIGN.CENTER
    t3_p.font.name = "Aptos"
    t3_p.font.size = Pt(10.5)
    t3_p.font.bold = True
    t3_p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    add_library_line(slide, lib_lines["side_fade_r"], col_lefts[2] + 10.0, col_top + 26.0, 270.0)

    o_base_y = col_top + 32.0
    o_step_y = 54.0
    for o_i, o_text in enumerate(d["opps_items"]):
        o_box = slide.shapes.add_textbox(Pt(col_lefts[2] + 10.0), Pt(o_base_y + o_i * o_step_y), Pt(col_w - 20.0), Pt(50.0))
        o_tf = o_box.text_frame
        o_tf.word_wrap = True
        o_tf.margin_top = Pt(1.0)
        o_tf.margin_bottom = Pt(1.0)
        o_p = o_tf.paragraphs[0]
        o_p.text = ""
        if ":" in o_text:
            parts = o_text.split(":", 1)
            r1 = o_p.add_run()
            r1.text = "• " + parts[0] + ":"
            r1.font.name = "Aptos"
            r1.font.size = Pt(9.0)
            r1.font.bold = True
            r1.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

            r2 = o_p.add_run()
            r2.text = parts[1]
            r2.font.name = "Aptos"
            r2.font.size = Pt(9.0)
            r2.font.bold = False
            r2.font.color.rgb = hex_to_rgb("334155")
        else:
            r = o_p.add_run()
            r.text = "• " + o_text
            r.font.name = "Aptos"
            r.font.size = Pt(9.0)
            r.font.color.rgb = hex_to_rgb("334155")

        if o_i < 4:
            add_library_line(slide, lib_lines["side_fade_r"], col_lefts[2] + 15.0, o_base_y + (o_i + 1) * o_step_y - 2.0, col_w - 30.0)

    # 4. Bottom Key Outcomes Bar
    out_bar = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(430.0), Pt(900.0), Pt(68.0))
    apply_solid_fill(out_bar, "FFF8F2")
    out_bar.line.color.rgb = hex_to_rgb(ColorTokens.VIVID_TANGERINE_MID)
    set_shape_corner_radius(out_bar, 4.0)
    apply_library_shadow(out_bar)

    # Left badge
    badge = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(42.0), Pt(442.0), Pt(115.0), Pt(44.0))
    apply_solid_fill(badge, ColorTokens.BRAND_INDIGO)
    badge.line.fill.background()
    set_shape_corner_radius(badge, 4.0)
    b_p = badge.text_frame.paragraphs[0]
    b_p.text = d["outcomes_header"]
    b_p.alignment = pptx.enum.text.PP_ALIGN.CENTER
    b_p.font.name = "Aptos"
    b_p.font.size = Pt(11.0)
    b_p.font.bold = True
    b_p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # 4 outcome tiles
    out_widths = [185.0, 175.0, 175.0, 180.0]
    out_lefts = [168.0, 360.0, 542.0, 725.0]
    for m_i, m_txt in enumerate(d["outcomes_metrics"]):
        m_box = slide.shapes.add_textbox(Pt(out_lefts[m_i]), Pt(438.0), Pt(out_widths[m_i]), Pt(52.0))
        m_tf = m_box.text_frame
        m_tf.word_wrap = True
        m_tf.margin_top = Pt(2.0)
        m_tf.margin_bottom = Pt(2.0)
        m_p = m_tf.paragraphs[0]
        m_p.text = m_txt
        m_p.alignment = pptx.enum.text.PP_ALIGN.CENTER
        m_p.font.name = "Aptos"
        m_p.font.size = Pt(9.0)
        m_p.font.bold = True
        if m_i == 0 or m_i == 3:
            m_p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_TANGERINE)
        else:
            m_p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

        if m_i < 3:
            add_library_line(slide, lib_lines["solid_v"], out_lefts[m_i] + out_widths[m_i] + 4.0, 444.0, width_pt=0.0, height_pt=42.0)


def build_slide5(prs, d):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # 1. Title
    title_box = slide.shapes.add_textbox(Pt(30.0), Pt(10.0), Pt(900.0), Pt(38.0))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = d["title"]
    p.font.name = "Aptos"
    p.font.size = Pt(14.5)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

    # 2. Table: 8 rows x 4 cols
    table_shape = slide.shapes.add_table(8, 4, Pt(30.0), Pt(54.0), Pt(900.0), Pt(440.0))
    tbl = table_shape.table
    tbl.columns[0].width = Pt(170.0)
    tbl.columns[1].width = Pt(120.0)
    tbl.columns[2].width = Pt(110.0)
    tbl.columns[3].width = Pt(500.0)

    # Header Row 0
    for c_idx, h_text in enumerate(d["headers"]):
        cell = tbl.cell(0, c_idx)
        cell.fill.solid()
        r = int(ColorTokens.BRAND_INDIGO[0:2], 16)
        g = int(ColorTokens.BRAND_INDIGO[2:4], 16)
        b = int(ColorTokens.BRAND_INDIGO[4:6], 16)
        cell.fill.fore_color.rgb = RGBColor(r, g, b)
        set_cell_border(cell, color="CBD5E1", width="6350")

        tf = cell.text_frame
        tf.word_wrap = True
        tf.margin_top = Pt(4.0)
        tf.margin_bottom = Pt(4.0)
        p = tf.paragraphs[0]
        p.text = h_text
        p.font.name = "Aptos"
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        if c_idx == 1 or c_idx == 2:
            p.alignment = pptx.enum.text.PP_ALIGN.CENTER
        else:
            p.alignment = pptx.enum.text.PP_ALIGN.LEFT

    # Data Rows 1 to 7
    for r_idx, row_vals in enumerate(d["rows"]):
        bg_hex = "F8FAFC" if (r_idx % 2 == 1) else "FFFFFF"
        for c_idx, cell_paras in enumerate(row_vals):
            cell = tbl.cell(r_idx + 1, c_idx)
            cell.fill.solid()
            r = int(bg_hex[0:2], 16)
            g = int(bg_hex[2:4], 16)
            b = int(bg_hex[4:6], 16)
            cell.fill.fore_color.rgb = RGBColor(r, g, b)
            set_cell_border(cell, color="E2E8F0", width="6350")

            tf = cell.text_frame
            tf.word_wrap = True
            tf.margin_top = Pt(2.5)
            tf.margin_bottom = Pt(2.5)
            tf.margin_left = Pt(6.0)
            tf.margin_right = Pt(6.0)

            for p_idx, p_text in enumerate(cell_paras):
                p = tf.paragraphs[0] if p_idx == 0 else tf.add_paragraph()
                p.text = p_text
                p.font.name = "Aptos"
                if c_idx == 0:
                    p.font.size = Pt(9.5)
                    p.font.bold = True
                    p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)
                    p.alignment = pptx.enum.text.PP_ALIGN.LEFT
                elif c_idx == 1:
                    p.font.size = Pt(11.0)
                    p.font.bold = True
                    p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)
                    p.alignment = pptx.enum.text.PP_ALIGN.CENTER
                elif c_idx == 2:
                    p.font.size = Pt(11.0)
                    p.font.bold = True
                    p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)
                    p.alignment = pptx.enum.text.PP_ALIGN.CENTER
                else:
                    p.font.size = Pt(9.5)
                    p.font.color.rgb = hex_to_rgb("334155")
                    p.alignment = pptx.enum.text.PP_ALIGN.LEFT


def build_slide6(prs, d, lib_lines):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # 1. Title
    title_box = slide.shapes.add_textbox(Pt(30.0), Pt(10.0), Pt(900.0), Pt(36.0))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = d["title"]
    p.font.name = "Aptos"
    p.font.size = Pt(15.0)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

    # 2. Header Bar
    hdr_lefts = [30.0, 218.0, 396.0, 804.0]
    hdr_widths = [180.0, 170.0, 400.0, 126.0]
    for h_i, h_txt in enumerate(d["headers"]):
        h_tab = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(hdr_lefts[h_i]), Pt(62.7), Pt(hdr_widths[h_i]), Pt(29.0))
        apply_solid_fill(h_tab, ColorTokens.BRAND_INDIGO)
        h_tab.line.fill.background()
        set_shape_corner_radius(h_tab, 4.0)
        hp = h_tab.text_frame.paragraphs[0]
        hp.text = h_txt
        hp.alignment = pptx.enum.text.PP_ALIGN.CENTER
        hp.font.name = "Aptos"
        hp.font.size = Pt(12.0)
        hp.font.bold = True
        hp.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # 3. 4 Workstream Cards (A, B, C, D) - Tops at 100.7, 202.7, 304.7, 406.7 pt
    row_tops = [100.7, 202.7, 304.7, 406.7]
    row_h = 96.0

    for r_i, tr in enumerate(d["tracks"]):
        ry = row_tops[r_i]

        # Card background with Vivid Indigo Soft border (#B1A8D8) & Decor_Box_with_Shadow_1 shadow
        card = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(ry), Pt(900.0), Pt(row_h))
        apply_solid_fill(card, "FFFFFF")
        card.line.color.rgb = hex_to_rgb(ColorTokens.VIVID_INDIGO_SOFT)
        card.line.width = Pt(0.25)
        set_shape_corner_radius(card, 4.0)
        apply_library_shadow(card)

        # Col 1: Unified Badge (#B1A8D8 + shadow) + Centered Name (12pt)
        badge = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.OVAL, Pt(42.0), Pt(ry + 30.0), Pt(36.0), Pt(36.0))
        apply_solid_fill(badge, ColorTokens.VIVID_INDIGO_SOFT)
        badge.line.fill.background()
        apply_library_shadow(badge)
        bp = badge.text_frame.paragraphs[0]
        bp.text = tr["letter"]
        bp.alignment = pptx.enum.text.PP_ALIGN.CENTER
        bp.font.name = "Aptos"
        bp.font.size = Pt(16.0)
        bp.font.bold = True
        bp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

        # Name
        name_box = slide.shapes.add_textbox(Pt(84.0), Pt(ry + 29.8), Pt(124.0), Pt(36.4))
        ntf = name_box.text_frame
        ntf.word_wrap = True
        np = ntf.paragraphs[0]
        np.text = tr["name"]
        np.font.name = "Aptos"
        np.font.size = Pt(12.0)
        np.font.bold = True
        np.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)

        # Divider 1 (Authentic Grey Solid V from library, height=72, centered: y = ry + 12.0)
        add_library_line(slide, lib_lines["solid_v"], 214.0, ry + 12.0, width_pt=0.0, height_pt=72.0)

        # Col 2: Spend and Suppliers (Header: 11pt, Metric: 14pt)
        sp_box = slide.shapes.add_textbox(Pt(220.0), Pt(ry + 27.2), Pt(168.0), Pt(41.6))
        sptf = sp_box.text_frame
        sptf.word_wrap = True
        sp1 = sptf.paragraphs[0]
        sp1.text = tr["spend_sup"][0]
        sp1.font.name = "Aptos"
        sp1.font.size = Pt(11.0)
        sp1.font.bold = True
        sp1.font.color.rgb = hex_to_rgb("334155")

        sp2 = sptf.add_paragraph()
        sp2.text = tr["spend_sup"][1]
        sp2.font.name = "Aptos"
        sp2.font.size = Pt(14.0)
        sp2.font.bold = True
        sp2.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)
        sp2.space_before = Pt(3.0)

        # Divider 2 (Authentic Grey Solid V from library, height=72, centered: y = ry + 12.0)
        add_library_line(slide, lib_lines["solid_v"], 392.0, ry + 12.0, width_pt=0.0, height_pt=72.0)

        # Col 3: Planned Interventions (11pt bullets)
        pi_box = slide.shapes.add_textbox(Pt(400.0), Pt(ry + 22.3), Pt(394.0), Pt(51.3))
        pitf = pi_box.text_frame
        pitf.word_wrap = True
        for p_i, pi_txt in enumerate(tr["interventions"]):
            p = pitf.paragraphs[0] if p_i == 0 else pitf.add_paragraph()
            p.text = "• " + pi_txt
            p.font.name = "Aptos"
            p.font.size = Pt(11.0)
            p.font.color.rgb = hex_to_rgb("334155")
            p.space_after = Pt(2.0)

        # Divider 3 (Authentic Grey Solid V from library, height=72, centered: y = ry + 12.0)
        add_library_line(slide, lib_lines["solid_v"], 800.0, ry + 12.0, width_pt=0.0, height_pt=72.0)

        # Col 4: Estimated Impact (Top metric: 16pt Tangerine #FE860E, sub: 10.5pt)
        imp_box = slide.shapes.add_textbox(Pt(806.0), Pt(ry + 25.0), Pt(118.0), Pt(46.0))
        imptf = imp_box.text_frame
        imptf.word_wrap = True
        ip1 = imptf.paragraphs[0]
        ip1.text = tr["impact"][0]
        ip1.alignment = pptx.enum.text.PP_ALIGN.CENTER
        ip1.font.name = "Aptos"
        ip1.font.size = Pt(16.0)
        ip1.font.bold = True
        ip1.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_TANGERINE)

        ip2 = imptf.add_paragraph()
        ip2.text = tr["impact"][1]
        ip2.alignment = pptx.enum.text.PP_ALIGN.CENTER
        ip2.font.name = "Aptos"
        ip2.font.size = Pt(10.5)
        ip2.font.bold = True
        ip2.font.color.rgb = hex_to_rgb("#475569")


def main():
    print("Loading presentation...")
    # Refresh inspect copy from source (fallback to PowerShell if PowerPoint has file locked)
    src_path = BASE_DIR / "Design tester new.pptx"
    try:
        shutil.copyfile(src_path, INPUT_PATH)
    except Exception:
        import subprocess
        subprocess.run(
            ["powershell", "-NoProfile", "-Command", f"Copy-Item -LiteralPath '{src_path}' -Destination '{INPUT_PATH}' -Force"],
            capture_output=True,
            check=False,
        )
    
    in_prs = pptx.Presentation(INPUT_PATH)
    print(f"Loaded input presentation with {len(in_prs.slides)} slides.")

    s1_data = extract_slide1_data(in_prs)
    s2_data = extract_slide2_data(in_prs)
    s3_data = extract_slide3_data(in_prs)
    s4_data = extract_slide4_data(in_prs)
    s5_data = extract_slide5_data(in_prs)
    s6_data = extract_slide6_data(in_prs)

    print("Extracted all 6 slides successfully.")

    # Load Library prototypes and lines
    lib_prs, lib_s1, lib_s3, lib_lines = load_library()

    out_prs = pptx.Presentation()
    out_prs.slide_width = Pt(960)
    out_prs.slide_height = Pt(540)

    print("[Formatting Slide 1] 3-Column Floating Pill Cards...")
    build_slide1(out_prs, s1_data, lib_s1, lib_lines)

    print("[Formatting Slide 2] Split Process Diagnostics & Insights...")
    build_slide2(out_prs, s2_data, lib_s1, lib_s3, lib_lines)

    print("[Formatting Slide 3] Full-Width Process Rows...")
    build_slide3(out_prs, s3_data, lib_s3, lib_lines)

    print("[Formatting Slide 4] Case Study 3: Opportunity Assessment...")
    build_slide4(out_prs, s4_data, lib_lines)

    print("[Formatting Slide 5] QXO Leadership Priorities Table...")
    build_slide5(out_prs, s5_data)

    print("[Formatting Slide 6] Opportunity Perspective Workstreams...")
    build_slide6(out_prs, s6_data, lib_lines)

    print("Total slides built:", len(out_prs.slides))

    print("Enforcing exact 4px border radius on all rounded boxes across the deck...")
    enforce_deck_corner_radii(out_prs, 4.0)

    raw_buf = io.BytesIO()
    out_prs.save(raw_buf)
    raw_bytes = raw_buf.getvalue()

    print("Injecting SVG vector icons into Slide 2...")
    final_bytes = inject_svg_assets(raw_bytes, targets={2: "rIdSvgS2_"})

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_PATH, "wb") as f:
        f.write(final_bytes)
    with open(LOCAL_COPY_PATH, "wb") as f:
        f.write(final_bytes)

    print(f"Successfully generated formatted 6-slide presentations at:")
    print(f"  - {OUTPUT_PATH}")
    print(f"  - {LOCAL_COPY_PATH}")


if __name__ == "__main__":
    main()

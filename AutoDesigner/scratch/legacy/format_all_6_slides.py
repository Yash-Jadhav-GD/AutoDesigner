"""
AutoDesigner Precision Formatting Engine for All 6 Slides
Strictly adheres to:
1. Lines from Library_copy Slide 6:
   - Grey Side Fade R (horizontal fade right)
   - Grey Centre Fade H (horizontal center fade)
   - Grey Solid H (horizontal solid)
   - Grey Solid V (vertical solid)
   - Grey Centre Fade V (vertical center fade)
2. COLOR_THEORY_SPECIFICATION.md & After_copy.pptx visual benchmarks:
   - Slide 1: 3-Part transformation narrative (Problem -> Bridge -> Future)
   - Slide 2: Unified vivid.indigo.soft (#B1A8D8) backplates for parallel diagnostic points + Tangerine INSIGHTS callout
   - Slide 3: Full-width unified vivid.indigo.soft (#B1A8D8) backplates
   - Slide 4: Phase 1 Current State Assessment (5 diagnostic rows + Deliverables card + Diagnostic Insights KPI card with Grey Centre Fade H lines)
   - Slide 5: Phase 3 Multi-Agent Pilot Workflows (5 agent rows with 5-phase chromatic badges A-E & autonomy pills)
   - Slide 6: WS2 Team Structure (Leadership badge + 3x3 Persona grid cards with FTE badges + Total badge)
3. 100% Text Fidelity (ZERO added words, ZERO subtracted words)
4. Strict DrawingML schema compliance & clean OpenXML relationships
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
    inject_svg_assets,
    set_shape_corner_radius,
)
from engine.color_tokens import (
    ColorTokens,
    hex_to_rgb,
    apply_linear_gradient,
    apply_solid_fill,
)

INPUT_PATH = BASE_DIR / "Design tester new.pptx"
AFTER_PATH = BASE_DIR / "After_copy.pptx"
LIBRARY_PATH = BASE_DIR / "Library_copy.pptx"
LIBRARY_INSPECT_PATH = BASE_DIR / "Library_inspect.pptx"
OUTPUT_PATH = BASE_DIR / "output" / "Design_tester_new_formatted.pptx"
LOCAL_COPY_PATH = BASE_DIR / "Design tester new_formatted.pptx"








# ==============================================================================
# DATA EXTRACTION FUNCTIONS
# ==============================================================================

def extract_slide1_data(input_prs):
    s = input_prs.slides[0]
    shapes_by_x = {}
    for shape in s.shapes:
        if shape.has_text_frame and shape.text_frame.text.strip():
            x_col = round((shape.left / 12700) / 150)
            if x_col not in shapes_by_x:
                shapes_by_x[x_col] = []
            shapes_by_x[x_col].append(shape)

    columns = []
    for col_key in sorted(shapes_by_x.keys()):
        col_shapes = sorted(shapes_by_x[col_key], key=lambda sh: sh.top)
        header_text = col_shapes[0].text_frame.text.strip()
        body_texts = [sh.text_frame.text.strip() for sh in col_shapes[1:]]
        columns.append({"header": header_text, "items": body_texts})
    return columns


def extract_slide2_data(input_prs):
    s = input_prs.slides[1]
    left_shapes = []
    right_shapes = []

    for sh in s.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            if (sh.left / 12700) < 600.0:
                left_shapes.append(sh)
            else:
                right_shapes.append(sh)

    left_shapes.sort(key=lambda sh: sh.top)
    right_shapes.sort(key=lambda sh: sh.top)

    left_pairs = []
    for i in range(0, len(left_shapes), 2):
        sub_text = left_shapes[i].text_frame.text.strip()
        body_text = left_shapes[i + 1].text_frame.text.strip() if i + 1 < len(left_shapes) else ""
        left_pairs.append({"subheader": sub_text, "body": body_text})

    right_header = right_shapes[0].text_frame.text.strip() if right_shapes else "INSIGHTS"
    right_items = [sh.text_frame.text.strip() for sh in right_shapes[1:]]

    return {
        "left_pairs": left_pairs,
        "right_header": right_header,
        "right_items": right_items,
    }


def extract_slide3_data(input_prs):
    s = input_prs.slides[2]
    text_shapes = [sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip()]
    text_shapes.sort(key=lambda sh: sh.top)

    pairs = []
    for i in range(0, len(text_shapes), 2):
        sub_text = text_shapes[i].text_frame.text.strip()
        body_text = text_shapes[i + 1].text_frame.text.strip() if i + 1 < len(text_shapes) else ""
        pairs.append({"subheader": sub_text, "body": body_text})
    return pairs


def extract_slide4_data(input_prs):
    s = input_prs.slides[3]
    shapes = [sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip()]
    
    title = ""
    subtitle = ""
    left_shapes = []
    deliv_shapes = []
    insight_shapes = []

    for sh in shapes:
        x = sh.left / 12700
        y = sh.top / 12700
        txt = sh.text_frame.text.strip()
        if y < 35.0:
            title = txt
        elif y < 70.0:
            subtitle = txt
        elif x < 500.0:
            left_shapes.append(sh)
        else:
            if y < 240.0:
                deliv_shapes.append(sh)
            else:
                insight_shapes.append(sh)

    left_shapes.sort(key=lambda sh: sh.top)
    deliv_shapes.sort(key=lambda sh: sh.top)
    insight_shapes.sort(key=lambda sh: sh.top)

    left_pairs = []
    for i in range(0, len(left_shapes), 2):
        sub_text = left_shapes[i].text_frame.text.strip()
        body_text = left_shapes[i + 1].text_frame.text.strip() if i + 1 < len(left_shapes) else ""
        left_pairs.append({"subheader": sub_text, "body": body_text})

    deliv_header = deliv_shapes[0].text_frame.text.strip() if deliv_shapes else "PHASE 1 DELIVERABLES"
    deliv_items = [sh.text_frame.text.strip() for sh in deliv_shapes[1:]]

    insights_header = insight_shapes[0].text_frame.text.strip() if insight_shapes else "KEY DIAGNOSTIC INSIGHTS"
    kpi_shapes = insight_shapes[1:]
    kpi_rows = []
    for i in range(0, len(kpi_shapes), 3):
        val = kpi_shapes[i].text_frame.text.strip()
        lbl = kpi_shapes[i+1].text_frame.text.strip() if i+1 < len(kpi_shapes) else ""
        sub = kpi_shapes[i+2].text_frame.text.strip() if i+2 < len(kpi_shapes) else ""
        kpi_rows.append({"val": val, "label": lbl, "sub": sub})

    return {
        "title": title,
        "subtitle": subtitle,
        "left_pairs": left_pairs,
        "deliv_header": deliv_header,
        "deliv_items": deliv_items,
        "insights_header": insights_header,
        "kpi_rows": kpi_rows,
    }


def extract_slide5_data(input_prs):
    s = input_prs.slides[4]
    shapes = [sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip()]
    
    title = ""
    subtitle = ""
    row_shapes = []

    for sh in shapes:
        y = sh.top / 12700
        txt = sh.text_frame.text.strip()
        if y < 35.0:
            title = txt
        elif y < 70.0:
            subtitle = txt
        else:
            row_shapes.append(sh)

    row_shapes.sort(key=lambda sh: sh.top)
    rows_grouped = []
    current_group = []
    last_top = -1

    for sh in row_shapes:
        y = sh.top / 12700
        if last_top < 0 or abs(y - last_top) < 65.0:
            current_group.append(sh)
            last_top = y if last_top < 0 else last_top
        else:
            rows_grouped.append(current_group)
            current_group = [sh]
            last_top = y
    if current_group:
        rows_grouped.append(current_group)

    workflows = []
    for grp in rows_grouped:
        grp.sort(key=lambda sh: sh.left)
        letter = ""
        agent_title = ""
        desc = ""
        in_param = ""
        out_param = ""
        autonomy = ""
        for sh in grp:
            t = sh.text_frame.text.strip()
            if len(t) == 1 and t in "ABCDE":
                letter = t
            elif "%" in t or ("autonomous" in t.lower() and "Agent" not in t):
                autonomy = t
            elif t.startswith("In:"):
                in_param = t
            elif t.startswith("Out:"):
                out_param = t
            elif "Agent" in t:
                agent_title = t
            else:
                desc = t
        workflows.append({
            "letter": letter,
            "title": agent_title,
            "desc": desc,
            "in": in_param,
            "out": out_param,
            "autonomy": autonomy,
        })

    return {
        "title": title,
        "subtitle": subtitle,
        "workflows": workflows,
    }


def extract_slide6_data(input_prs):
    s = input_prs.slides[5]
    shapes = [sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip()]

    title = ""
    subtitle = ""
    lead_badge = ""
    total_badge = ""

    for sh in shapes:
        y = sh.top / 12700
        txt = sh.text_frame.text.strip()
        if y < 35.0:
            title = txt
        elif y < 70.0:
            subtitle = txt
        elif "Process Design Lead" in txt:
            lead_badge = txt
        elif "TOTAL:" in txt:
            total_badge = txt

    card_shapes = [sh for sh in shapes if sh.text_frame.text.strip() not in [title, subtitle, lead_badge, total_badge]]
    grid = {}
    for sh in card_shapes:
        x = sh.left / 12700
        y = sh.top / 12700
        c_idx = 0 if x < 330 else (1 if x < 630 else 2)
        r_idx = 0 if y < 250 else (1 if y < 360 else 2)
        key = (r_idx, c_idx)
        if key not in grid:
            grid[key] = []
        grid[key].append(sh)

    persona_cards = []
    for r in range(3):
        for c in range(3):
            items = grid.get((r, c), [])
            count_txt = ""
            label_txt = ""
            name_txt = ""
            region_txt = ""
            for sh in items:
                t = sh.text_frame.text.strip()
                if t.isdigit():
                    count_txt = t
                elif t in ["FTE", "FTEs"]:
                    label_txt = t
                elif any(delim in t for delim in ["·", "NA", "GC", "India", "LATAM", "EU"]) and not any(title_word in t for title_word in ["Digitalists", "Logistics", "Marketing", "COTBS", "VMO", "S&I"]):
                    region_txt = t
                else:
                    name_txt = t
            persona_cards.append({
                "count": count_txt,
                "label": label_txt,
                "name": name_txt,
                "regions": region_txt,
            })

    return {
        "title": title,
        "subtitle": subtitle,
        "lead_badge": lead_badge,
        "total_badge": total_badge,
        "persona_cards": persona_cards,
    }


# ==============================================================================
# SLIDE BUILDERS (1 - 6) WITH AUTHENTIC LIBRARY LINES
# ==============================================================================

def build_slide1(prs, s1_data, lib_s1, lib_lines):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    d1_box_proto = get_shape_by_name(lib_s1, "Design 1 Decor box")
    d1_hdr_proto = get_shape_by_name(lib_s1, "Design 1 Header Text")
    d1_body_proto = get_shape_by_name(lib_s1, "Design 1 Body text 1")

    column_x = [30.2, 341.8, 653.5]
    card_w = 276.4
    card_top = 80.0
    card_h = 410.0

    pill_styles = [
        {"start": ColorTokens.GRADIENT_INDIGO_START, "end": ColorTokens.GRADIENT_INDIGO_END, "text_color": hex_to_rgb(ColorTokens.TEXT_ON_DARK)},
        {"start": ColorTokens.GRADIENT_NAVY_START, "end": ColorTokens.GRADIENT_NAVY_END, "text_color": hex_to_rgb(ColorTokens.TEXT_ON_DARK)},
        {"start": ColorTokens.GRADIENT_TANGERINE_START, "end": ColorTokens.GRADIENT_TANGERINE_END, "text_color": hex_to_rgb(ColorTokens.TEXT_ON_DARK)},
    ]

    for col_idx, col_data in enumerate(s1_data):
        bx = column_x[col_idx]
        style = pill_styles[col_idx]

        elem = clean_element(copy.deepcopy(d1_box_proto._element))
        slide.shapes._spTree.append(elem)
        box = slide.shapes[len(slide.shapes) - 1]
        box.left = Pt(bx)
        box.top = Pt(card_top)
        box.width = Pt(card_w)
        box.height = Pt(card_h)
        set_shape_corner_radius(box, 4.0)

        elem = clean_element(copy.deepcopy(d1_hdr_proto._element))
        slide.shapes._spTree.append(elem)
        hdr = slide.shapes[len(slide.shapes) - 1]
        hdr.left = Pt(bx + 30.5)
        hdr.top = Pt(card_top - 15.7)
        hdr.width = Pt(215.4)
        hdr.height = Pt(33.0)
        set_shape_corner_radius(hdr, 4.0)
        apply_linear_gradient(hdr, style["start"], style["end"], angle_dml=10800000, with_shadow=True)
        hdr.text_frame.text = col_data["header"]
        for p in hdr.text_frame.paragraphs:
            p.font.name = "Aptos"
            p.font.size = Pt(15)
            p.font.bold = True
            p.font.color.rgb = style["text_color"]

        padding = 8.5039
        start_y = card_top + 33.0 + padding
        end_y = card_top + card_h - padding
        available_space = end_y - start_y
        row_height = 56.0
        sum_text_heights = len(col_data["items"]) * row_height
        total_empty_space = available_space - sum_text_heights
        num_gaps = (len(col_data["items"]) - 1) * 2
        gap_height = max(4.0, total_empty_space / num_gaps)

        current_y = start_y
        body_left = bx + 12.0
        body_width = card_w - 24.0

        for item_idx, body_text in enumerate(col_data["items"]):
            elem = clean_element(copy.deepcopy(d1_body_proto._element))
            slide.shapes._spTree.append(elem)
            b_box = slide.shapes[len(slide.shapes) - 1]
            b_box.left = Pt(body_left)
            b_box.top = Pt(current_y)
            b_box.width = Pt(body_width)
            b_box.height = Pt(row_height)
            b_box.text_frame.text = body_text
            b_box.text_frame.word_wrap = True
            for p in b_box.text_frame.paragraphs:
                p.font.name = "Aptos"
                p.font.size = Pt(11)
                p.font.bold = False
                p.font.color.rgb = hex_to_rgb(ColorTokens.TEXT_PRIMARY)

            current_y += row_height

            if item_idx < len(col_data["items"]) - 1:
                current_y += gap_height
                # Authentic Library Line: Grey Side Fade R from Slide 6
                add_library_line(slide, lib_lines["side_fade_r"], bx + 17.5, current_y, card_w - 35.0)
                current_y += gap_height


def build_slide2(prs, s2_data, lib_s1, lib_s3, lib_lines):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    decor_proto = get_shape_by_name(lib_s3, "UI-ROW-01 Decor Box")
    icon_box_proto = get_shape_by_name(lib_s3, "UI-ROW-01 Icon Box")
    subheader_proto = get_shape_by_name(lib_s3, "UI-ROW-01 Sub Header Text")
    body_proto = get_shape_by_name(lib_s3, "UI-ROW-01 Body Text")

    d1_box_proto = get_shape_by_name(lib_s1, "Design 1 Decor box")
    d1_hdr_proto = get_shape_by_name(lib_s1, "Design 1 Header Text")
    d1_body_proto = get_shape_by_name(lib_s1, "Design 1 Body text 1")

    left_x = 30.2
    left_w = 590.0
    row_count = len(s2_data["left_pairs"])
    row_h = 72.0
    row_gap = 12.5
    start_top = 65.0
    diagnostic_color = ColorTokens.VIVID_INDIGO_SOFT

    for idx, item in enumerate(s2_data["left_pairs"]):
        current_top = start_top + idx * (row_h + row_gap)

        elem = clean_element(copy.deepcopy(decor_proto._element))
        slide.shapes._spTree.append(elem)
        box = slide.shapes[len(slide.shapes) - 1]
        box.left = Pt(left_x)
        box.top = Pt(current_top)
        box.width = Pt(left_w)
        box.height = Pt(row_h)
        set_shape_corner_radius(box, 4.0)

        elem = clean_element(copy.deepcopy(icon_box_proto._element))
        slide.shapes._spTree.append(elem)
        ibox = slide.shapes[len(slide.shapes) - 1]
        ibox.left = Pt(left_x + 10.0)
        ibox.top = Pt(current_top + 10.1)
        ibox.width = Pt(51.7)
        ibox.height = Pt(51.7)
        apply_solid_fill(ibox, diagnostic_color)
        ibox.text_frame.text = ""

        # Authentic SVG Icon centered inside circle
        icon_w_pt = 28.4
        icon_h_pt = 28.4
        icon_l_pt = left_x + 10.0 + (51.7 - icon_w_pt) / 2.0
        icon_t_pt = current_top + 10.1 + (51.7 - icon_h_pt) / 2.0
        emu_l = int(icon_l_pt * 12700)
        emu_t = int(icon_t_pt * 12700)
        emu_w = int(icon_w_pt * 12700)
        emu_h = int(icon_h_pt * 12700)
        pic_id = 8100 + idx
        rel_id = f"rIdSvgS2_{idx+1}"
        pic_xml = f'''<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:nvPicPr>
            <p:cNvPr id="{pic_id}" name="Graphic S2 {200+idx}"/>
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
            <a:xfrm><a:off x="{emu_l}" y="{emu_t}"/><a:ext cx="{emu_w}" cy="{emu_h}"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          </p:spPr>
        </p:pic>'''
        slide.shapes._spTree.append(parse_xml(pic_xml))

        elem = clean_element(copy.deepcopy(subheader_proto._element))
        slide.shapes._spTree.append(elem)
        sub = slide.shapes[len(slide.shapes) - 1]
        sub.left = Pt(left_x + 72.0)
        sub.top = Pt(current_top + 10.0)
        sub.width = Pt(left_w - 82.0)
        sub.height = Pt(18.0)
        sub.text_frame.text = item["subheader"]
        for p in sub.text_frame.paragraphs:
            p.font.name = "Aptos"
            p.font.size = Pt(13.5)
            p.font.bold = True
            p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)

        # Authentic Library Line: Grey Side Fade R from Slide 6
        add_library_line(slide, lib_lines["side_fade_r"], left_x + 72.0, current_top + 31.0, left_w - 82.0)

        elem = clean_element(copy.deepcopy(body_proto._element))
        slide.shapes._spTree.append(elem)
        body = slide.shapes[len(slide.shapes) - 1]
        body.left = Pt(left_x + 72.0)
        body.top = Pt(current_top + 36.0)
        body.width = Pt(left_w - 82.0)
        body.height = Pt(26.0)
        body.text_frame.text = item["body"]
        body.text_frame.word_wrap = True
        for p in body.text_frame.paragraphs:
            p.font.name = "Aptos"
            p.font.size = Pt(11)
            p.font.color.rgb = hex_to_rgb(ColorTokens.TEXT_PRIMARY)

    # Right INSIGHTS card
    right_x = 650.0
    right_w = 280.0
    card_top = 80.0
    card_h = 410.0

    elem = clean_element(copy.deepcopy(d1_box_proto._element))
    slide.shapes._spTree.append(elem)
    box = slide.shapes[len(slide.shapes) - 1]
    box.left = Pt(right_x)
    box.top = Pt(card_top)
    box.width = Pt(right_w)
    box.height = Pt(card_h)
    set_shape_corner_radius(box, 4.0)

    elem = clean_element(copy.deepcopy(d1_hdr_proto._element))
    slide.shapes._spTree.append(elem)
    hdr = slide.shapes[len(slide.shapes) - 1]
    hdr.left = Pt(right_x + (right_w - 215.4) / 2.0)
    hdr.top = Pt(card_top - 15.7)
    hdr.width = Pt(215.4)
    hdr.height = Pt(33.0)
    set_shape_corner_radius(hdr, 4.0)
    apply_linear_gradient(hdr, ColorTokens.GRADIENT_TANGERINE_START, ColorTokens.GRADIENT_TANGERINE_END, angle_dml=10800000, with_shadow=True)
    hdr.text_frame.text = s2_data["right_header"]
    for p in hdr.text_frame.paragraphs:
        p.font.name = "Aptos"
        p.font.size = Pt(15)
        p.font.bold = True
        p.font.color.rgb = hex_to_rgb(ColorTokens.TEXT_ON_DARK)

    padding = 8.5039
    start_y = card_top + 33.0 + padding
    end_y = card_top + card_h - padding
    available_space = end_y - start_y
    row_height = 56.0
    sum_text_heights = len(s2_data["right_items"]) * row_height
    total_empty_space = available_space - sum_text_heights
    num_gaps = (len(s2_data["right_items"]) - 1) * 2
    gap_height = max(4.0, total_empty_space / num_gaps)

    current_y = start_y
    body_left = right_x + 12.0
    body_width = right_w - 24.0

    for item_idx, body_text in enumerate(s2_data["right_items"]):
        elem = clean_element(copy.deepcopy(d1_body_proto._element))
        slide.shapes._spTree.append(elem)
        b_box = slide.shapes[len(slide.shapes) - 1]
        b_box.left = Pt(body_left)
        b_box.top = Pt(current_y)
        b_box.width = Pt(body_width)
        b_box.height = Pt(row_height)
        b_box.text_frame.text = body_text
        b_box.text_frame.word_wrap = True
        for p in b_box.text_frame.paragraphs:
            p.font.name = "Aptos"
            p.font.size = Pt(11)
            p.font.bold = False
            p.font.color.rgb = hex_to_rgb(ColorTokens.TEXT_PRIMARY)

        current_y += row_height

        if item_idx < len(s2_data["right_items"]) - 1:
            current_y += gap_height
            # Authentic Library Line: Grey Side Fade R from Slide 6
            add_library_line(slide, lib_lines["side_fade_r"], right_x + 17.5, current_y, right_w - 35.0)
            current_y += gap_height


def build_slide3(prs, s3_data, lib_s3, lib_lines):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    decor_proto = get_shape_by_name(lib_s3, "UI-ROW-01 Decor Box")
    icon_box_proto = get_shape_by_name(lib_s3, "UI-ROW-01 Icon Box")
    subheader_proto = get_shape_by_name(lib_s3, "UI-ROW-01 Sub Header Text")
    body_proto = get_shape_by_name(lib_s3, "UI-ROW-01 Body Text")

    row_count = len(s3_data)
    row_h = 72.0
    row_gap = 14.0
    total_h = row_count * row_h + (row_count - 1) * row_gap
    start_top = (540.0 - total_h) / 2.0
    row_left = 24.8
    row_width = 910.4
    diagnostic_color = ColorTokens.VIVID_INDIGO_SOFT

    for idx, item in enumerate(s3_data):
        current_top = start_top + idx * (row_h + row_gap)

        elem = clean_element(copy.deepcopy(decor_proto._element))
        slide.shapes._spTree.append(elem)
        box = slide.shapes[len(slide.shapes) - 1]
        box.left = Pt(row_left)
        box.top = Pt(current_top)
        box.width = Pt(row_width)
        box.height = Pt(row_h)
        set_shape_corner_radius(box, 4.0)

        elem = clean_element(copy.deepcopy(icon_box_proto._element))
        slide.shapes._spTree.append(elem)
        ibox = slide.shapes[len(slide.shapes) - 1]
        ibox.left = Pt(row_left + 10.7)
        ibox.top = Pt(current_top + 10.1)
        ibox.width = Pt(51.7)
        ibox.height = Pt(51.7)
        apply_solid_fill(ibox, diagnostic_color)
        ibox.text_frame.text = ""

        elem = clean_element(copy.deepcopy(subheader_proto._element))
        slide.shapes._spTree.append(elem)
        sub = slide.shapes[len(slide.shapes) - 1]
        sub.left = Pt(row_left + 74.0)
        sub.top = Pt(current_top + 10.0)
        sub.width = Pt(row_width - 86.0)
        sub.height = Pt(18.0)
        sub.text_frame.text = item["subheader"]
        for p in sub.text_frame.paragraphs:
            p.font.name = "Aptos"
            p.font.size = Pt(13.5)
            p.font.bold = True
            p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)

        # Authentic Library Line: Grey Side Fade R from Slide 6
        add_library_line(slide, lib_lines["side_fade_r"], row_left + 74.0, current_top + 31.0, row_width - 86.0)

        elem = clean_element(copy.deepcopy(body_proto._element))
        slide.shapes._spTree.append(elem)
        body = slide.shapes[len(slide.shapes) - 1]
        body.left = Pt(row_left + 74.0)
        body.top = Pt(current_top + 36.0)
        body.width = Pt(row_width - 86.0)
        body.height = Pt(26.0)
        body.text_frame.text = item["body"]
        body.text_frame.word_wrap = True
        for p in body.text_frame.paragraphs:
            p.font.name = "Aptos"
            p.font.size = Pt(11)
            p.font.color.rgb = hex_to_rgb(ColorTokens.TEXT_PRIMARY)


def build_slide4(prs, s4_data, lib_s1, lib_s3, lib_lines):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # 1. Slide Title & Subtitle
    title_box = slide.shapes.add_textbox(Pt(20.2), Pt(14.0), Pt(900.0), Pt(48.0))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = s4_data["title"]
    p.font.name = "Aptos"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

    sub_box = slide.shapes.add_textbox(Pt(30.2), Pt(68.0), Pt(900.0), Pt(22.0))
    tf2 = sub_box.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = s4_data["subtitle"]
    p2.font.name = "Aptos"
    p2.font.size = Pt(11.5)
    p2.font.color.rgb = hex_to_rgb(ColorTokens.TEXT_SECONDARY)

    decor_proto = get_shape_by_name(lib_s3, "UI-ROW-01 Decor Box")
    icon_box_proto = get_shape_by_name(lib_s3, "UI-ROW-01 Icon Box")
    subheader_proto = get_shape_by_name(lib_s3, "UI-ROW-01 Sub Header Text")
    body_proto = get_shape_by_name(lib_s3, "UI-ROW-01 Body Text")

    # 2. Left 5 Process Rows (top: 102.2 to 506.0 pt)
    left_x = 30.2
    left_w = 528.0
    start_top = 102.2
    row_h = 71.1
    row_gap = 12.1
    circle_d = 42.8

    for idx, item in enumerate(s4_data["left_pairs"]):
        current_top = start_top + idx * (row_h + row_gap)

        elem = clean_element(copy.deepcopy(decor_proto._element))
        slide.shapes._spTree.append(elem)
        box = slide.shapes[len(slide.shapes) - 1]
        box.left = Pt(left_x)
        box.top = Pt(current_top)
        box.width = Pt(left_w)
        box.height = Pt(row_h)
        set_shape_corner_radius(box, 4.0)

        # Periwinkle circle backplate
        elem = clean_element(copy.deepcopy(icon_box_proto._element))
        slide.shapes._spTree.append(elem)
        ibox = slide.shapes[len(slide.shapes) - 1]
        ibox.left = Pt(left_x + 16.0)
        ibox.top = Pt(current_top + (row_h - circle_d) / 2.0)
        ibox.width = Pt(circle_d)
        ibox.height = Pt(circle_d)
        apply_solid_fill(ibox, ColorTokens.VIVID_INDIGO_SOFT)
        ibox.text_frame.text = ""

        # Authentic SVG Icon centered inside circle
        icon_w_pt = 28.4
        icon_h_pt = 28.4
        icon_l_pt = left_x + 16.0 + (circle_d - icon_w_pt) / 2.0
        icon_t_pt = current_top + (row_h - icon_h_pt) / 2.0
        emu_l = int(icon_l_pt * 12700)
        emu_t = int(icon_t_pt * 12700)
        emu_w = int(icon_w_pt * 12700)
        emu_h = int(icon_h_pt * 12700)
        pic_id = 9100 + idx
        rel_id = f"rIdSvgS4_{idx+1}"
        pic_xml = f'''<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:nvPicPr>
            <p:cNvPr id="{pic_id}" name="Graphic S4 {200+idx}"/>
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
            <a:xfrm><a:off x="{emu_l}" y="{emu_t}"/><a:ext cx="{emu_w}" cy="{emu_h}"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          </p:spPr>
        </p:pic>'''
        slide.shapes._spTree.append(parse_xml(pic_xml))

        elem = clean_element(copy.deepcopy(subheader_proto._element))
        slide.shapes._spTree.append(elem)
        sub = slide.shapes[len(slide.shapes) - 1]
        sub.left = Pt(left_x + 72.0)
        sub.top = Pt(current_top + 8.5)
        sub.width = Pt(left_w - 82.0)
        sub.height = Pt(18.0)
        sub.text_frame.text = item["subheader"]
        for p in sub.text_frame.paragraphs:
            p.font.name = "Aptos"
            p.font.size = Pt(13)
            p.font.bold = True
            p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)

        # Authentic Library Line: Grey Side Fade R from Slide 6
        add_library_line(slide, lib_lines["side_fade_r"], left_x + 72.0, current_top + 28.5, left_w - 82.0)

        elem = clean_element(copy.deepcopy(body_proto._element))
        slide.shapes._spTree.append(elem)
        body = slide.shapes[len(slide.shapes) - 1]
        body.left = Pt(left_x + 72.0)
        body.top = Pt(current_top + 33.5)
        body.width = Pt(left_w - 82.0)
        body.height = Pt(28.0)
        body.text_frame.text = item["body"]
        body.text_frame.word_wrap = True
        for p in body.text_frame.paragraphs:
            p.font.name = "Aptos"
            p.font.size = Pt(10)
            p.font.color.rgb = hex_to_rgb(ColorTokens.TEXT_PRIMARY)

    # 3. Right Top Card: PHASE 1 DELIVERABLES (Solid Indigo #7C6DB2 with Dark Indigo Tab & White text)
    right_x = 574.4
    right_w = 355.4
    deliv_top = 98.5
    deliv_h = 164.0

    d_box = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(right_x), Pt(deliv_top), Pt(right_w), Pt(deliv_h))
    apply_solid_fill(d_box, ColorTokens.VIVID_INDIGO_MID)
    d_box.line.fill.background()

    # Rounded top header tab
    hdr_shape = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUND_2_SAME_RECTANGLE, Pt(right_x + (right_w - 210.0) / 2.0), Pt(deliv_top), Pt(210.0), Pt(26.0))
    apply_solid_fill(hdr_shape, ColorTokens.BRAND_INDIGO)
    hdr_shape.line.fill.background()
    tf_dh = hdr_shape.text_frame
    p_dh = tf_dh.paragraphs[0]
    p_dh.text = s4_data["deliv_header"]
    p_dh.alignment = pptx.enum.text.PP_ALIGN.CENTER
    p_dh.font.name = "Aptos"
    p_dh.font.size = Pt(11)
    p_dh.font.bold = True
    p_dh.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # 5 Deliverable checkmark items in white
    item_y = deliv_top + 32.0
    for d_text in s4_data["deliv_items"]:
        d_item = slide.shapes.add_textbox(Pt(right_x + 18.0), Pt(item_y), Pt(right_w - 36.0), Pt(22.0))
        d_tf = d_item.text_frame
        d_tf.word_wrap = True
        d_p = d_tf.paragraphs[0]
        d_p.text = d_text
        d_p.font.name = "Aptos"
        d_p.font.size = Pt(10.5)
        d_p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        item_y += 24.5

    # 4. Right Bottom Card: KEY DIAGNOSTIC INSIGHTS (Tangerine border, Tangerine focal metrics, aligned to 506.0 pt)
    ins_top = 270.0
    ins_h = 236.0

    ins_box = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(right_x), Pt(ins_top), Pt(right_w), Pt(ins_h))
    apply_solid_fill(ins_box, ColorTokens.NEUTRAL_WHITE)
    ins_box.line.color.rgb = hex_to_rgb(ColorTokens.VIVID_TANGERINE_MID)
    ins_box.line.width = Pt(1.0)

    # Header title text inside card
    hdr_box = slide.shapes.add_textbox(Pt(right_x), Pt(ins_top + 8.0), Pt(right_w), Pt(24.0))
    tf_ih = hdr_box.text_frame
    p_ih = tf_ih.paragraphs[0]
    p_ih.text = s4_data["insights_header"]
    p_ih.alignment = pptx.enum.text.PP_ALIGN.CENTER
    p_ih.font.name = "Aptos"
    p_ih.font.size = Pt(11)
    p_ih.font.bold = True
    p_ih.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

    # 4 KPI metric rows with Tangerine values and right alignment
    kpi_start_y = ins_top + 38.0
    kpi_step = 49.0
    for k_idx, kpi in enumerate(s4_data["kpi_rows"]):
        current_kpi_y = kpi_start_y + k_idx * kpi_step

        val_box = slide.shapes.add_textbox(Pt(right_x + 14.0), Pt(current_kpi_y), Pt(82.0), Pt(32.0))
        val_tf = val_box.text_frame
        val_p = val_tf.paragraphs[0]
        val_p.text = kpi["val"]
        val_p.alignment = pptx.enum.text.PP_ALIGN.RIGHT
        val_p.font.name = "Aptos"
        val_p.font.size = Pt(18)
        val_p.font.bold = True
        val_p.font.color.rgb = hex_to_rgb(ColorTokens.VIVID_TANGERINE_MID)

        lbl_box = slide.shapes.add_textbox(Pt(right_x + 104.0), Pt(current_kpi_y - 2.0), Pt(right_w - 116.0), Pt(20.0))
        lbl_tf = lbl_box.text_frame
        lbl_p = lbl_tf.paragraphs[0]
        lbl_p.text = kpi["label"]
        lbl_p.font.name = "Aptos"
        lbl_p.font.size = Pt(10.5)
        lbl_p.font.bold = True
        lbl_p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

        sub_box = slide.shapes.add_textbox(Pt(right_x + 104.0), Pt(current_kpi_y + 16.0), Pt(right_w - 116.0), Pt(18.0))
        sub_tf = sub_box.text_frame
        sub_p = sub_tf.paragraphs[0]
        sub_p.text = kpi["sub"]
        sub_p.font.name = "Aptos"
        sub_p.font.size = Pt(9.5)
        sub_p.font.color.rgb = hex_to_rgb(ColorTokens.TEXT_SECONDARY)

        if k_idx < len(s4_data["kpi_rows"]) - 1:
            # Authentic Library Line: Grey Centre Fade H from Slide 6
            add_library_line(slide, lib_lines["centre_fade_h"], right_x + 16.0, current_kpi_y + 42.0, right_w - 32.0)


def build_slide5(prs, s5_data, lib_s3, lib_lines):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    title_box = slide.shapes.add_textbox(Pt(30.2), Pt(14.0), Pt(900.0), Pt(32.0))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = s5_data["title"]
    p.font.name = "Aptos"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

    sub_box = slide.shapes.add_textbox(Pt(30.2), Pt(46.0), Pt(900.0), Pt(22.0))
    tf2 = sub_box.text_frame
    p2 = tf2.paragraphs[0]
    p2.text = s5_data["subtitle"]
    p2.font.name = "Aptos"
    p2.font.size = Pt(11.5)
    p2.font.color.rgb = hex_to_rgb(ColorTokens.TEXT_SECONDARY)

    decor_proto = get_shape_by_name(lib_s3, "UI-ROW-01 Decor Box")

    row_x = 26.9
    row_w = 906.2
    start_top = 78.0
    row_h = 76.0
    row_gap = 10.0

    phase_colors = [
        {"bg": ColorTokens.VIVID_TANGERINE_MID, "fg": ColorTokens.BRAND_ONYX},
        {"bg": ColorTokens.VIVID_TANGERINE_SOFT, "fg": ColorTokens.BRAND_ONYX},
        {"bg": ColorTokens.VIVID_INDIGO_MID, "fg": ColorTokens.TEXT_ON_DARK},
        {"bg": ColorTokens.VIVID_INDIGO_SOFT, "fg": ColorTokens.BRAND_INDIGO},
        {"bg": ColorTokens.VIVID_NAVY_SOFT, "fg": ColorTokens.TEXT_ON_DARK},
    ]

    for idx, wf in enumerate(s5_data["workflows"]):
        current_top = start_top + idx * (row_h + row_gap)
        color_scheme = phase_colors[idx % len(phase_colors)]

        elem = clean_element(copy.deepcopy(decor_proto._element))
        slide.shapes._spTree.append(elem)
        box = slide.shapes[len(slide.shapes) - 1]
        box.left = Pt(row_x)
        box.top = Pt(current_top)
        box.width = Pt(row_w)
        box.height = Pt(row_h)

        badge = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.OVAL, Pt(row_x + 14.0), Pt(current_top + 13.0), Pt(50.0), Pt(50.0))
        apply_solid_fill(badge, color_scheme["bg"])
        badge.line.fill.background()
        b_tf = badge.text_frame
        b_p = b_tf.paragraphs[0]
        b_p.text = wf["letter"]
        b_p.alignment = pptx.enum.text.PP_ALIGN.CENTER
        b_p.font.name = "Aptos"
        b_p.font.size = Pt(16)
        b_p.font.bold = True
        b_p.font.color.rgb = hex_to_rgb(color_scheme["fg"])

        t_box = slide.shapes.add_textbox(Pt(row_x + 78.0), Pt(current_top + 8.0), Pt(360.0), Pt(22.0))
        t_tf = t_box.text_frame
        t_p = t_tf.paragraphs[0]
        t_p.text = wf["title"]
        t_p.font.name = "Aptos"
        t_p.font.size = Pt(13)
        t_p.font.bold = True
        t_p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)

        d_box = slide.shapes.add_textbox(Pt(row_x + 78.0), Pt(current_top + 28.0), Pt(360.0), Pt(42.0))
        d_tf = d_box.text_frame
        d_tf.word_wrap = True
        d_p = d_tf.paragraphs[0]
        d_p.text = wf["desc"]
        d_p.font.name = "Aptos"
        d_p.font.size = Pt(10)
        d_p.font.color.rgb = hex_to_rgb(ColorTokens.TEXT_PRIMARY)

        # In Parameter Textbox
        in_box = slide.shapes.add_textbox(Pt(row_x + 450.0), Pt(current_top + 12.0), Pt(270.0), Pt(24.0))
        in_tf = in_box.text_frame
        in_p = in_tf.paragraphs[0]
        in_p.text = wf["in"]
        in_p.font.name = "Aptos"
        in_p.font.size = Pt(9.5)
        in_p.font.color.rgb = hex_to_rgb(ColorTokens.TEXT_SECONDARY)

        # Out Parameter Textbox
        out_box = slide.shapes.add_textbox(Pt(row_x + 450.0), Pt(current_top + 38.0), Pt(270.0), Pt(24.0))
        out_tf = out_box.text_frame
        out_p = out_tf.paragraphs[0]
        out_p.text = wf["out"]
        out_p.font.name = "Aptos"
        out_p.font.size = Pt(9.5)
        out_p.font.color.rgb = hex_to_rgb(ColorTokens.TEXT_SECONDARY)

        # Autonomy Badge Pill
        auto_pill = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(row_x + 740.0), Pt(current_top + 20.0), Pt(145.0), Pt(34.0))
        apply_solid_fill(auto_pill, color_scheme["bg"])
        auto_pill.line.fill.background()
        a_tf = auto_pill.text_frame
        a_p = a_tf.paragraphs[0]
        a_p.text = wf["autonomy"]
        a_p.alignment = pptx.enum.text.PP_ALIGN.CENTER
        a_p.font.name = "Aptos"
        a_p.font.size = Pt(11)
        a_p.font.bold = True
        a_p.font.color.rgb = hex_to_rgb(color_scheme["fg"])


def build_slide6(prs, s6_data, lib_s3, lib_lines):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    title_box = slide.shapes.add_textbox(Pt(30.2), Pt(14.0), Pt(900.0), Pt(32.0))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = s6_data["title"]
    p.font.name = "Aptos"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

    sub_box = slide.shapes.add_textbox(Pt(30.2), Pt(46.0), Pt(900.0), Pt(22.0))
    tf2 = sub_box.text_frame
    p2 = tf2.paragraphs[0]
    p2.text = s6_data["subtitle"]
    p2.font.name = "Aptos"
    p2.font.size = Pt(11.5)
    p2.font.color.rgb = hex_to_rgb(ColorTokens.TEXT_SECONDARY)

    panel = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(22.7), Pt(78.0), Pt(914.6), Pt(438.0))
    apply_solid_fill(panel, ColorTokens.NEUTRAL_SURFACE)
    panel.line.color.rgb = hex_to_rgb(ColorTokens.NEUTRAL_BORDER)
    panel.line.width = Pt(0.75)

    lead_shape = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(221.1), Pt(65.0), Pt(517.8), Pt(32.0))
    apply_solid_fill(lead_shape, ColorTokens.BRAND_INDIGO)
    lead_shape.line.fill.background()
    l_tf = lead_shape.text_frame
    l_p = l_tf.paragraphs[0]
    l_p.text = s6_data["lead_badge"]
    l_p.alignment = pptx.enum.text.PP_ALIGN.CENTER
    l_p.font.name = "Aptos"
    l_p.font.size = Pt(11)
    l_p.font.bold = True
    l_p.font.color.rgb = hex_to_rgb(ColorTokens.TEXT_ON_DARK)

    col_x = [46.0, 344.0, 642.0]
    row_y = [112.0, 222.0, 332.0]
    card_w = 272.0
    card_h = 92.0

    decor_proto = get_shape_by_name(lib_s3, "UI-ROW-01 Decor Box")

    for idx, card in enumerate(s6_data["persona_cards"]):
        r = idx // 3
        c = idx % 3
        bx = col_x[c]
        by = row_y[r]

        elem = clean_element(copy.deepcopy(decor_proto._element))
        slide.shapes._spTree.append(elem)
        box = slide.shapes[len(slide.shapes) - 1]
        box.left = Pt(bx)
        box.top = Pt(by)
        box.width = Pt(card_w)
        box.height = Pt(card_h)

        badge = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(bx + 8.0), Pt(by + 8.0), Pt(64.0), Pt(76.0))
        apply_solid_fill(badge, ColorTokens.VIVID_INDIGO_SOFT)
        badge.line.fill.background()
        badge.text_frame.text = ""

        cnt_box = slide.shapes.add_textbox(Pt(bx + 8.0), Pt(by + 10.0), Pt(64.0), Pt(38.0))
        cnt_tf = cnt_box.text_frame
        cnt_p = cnt_tf.paragraphs[0]
        cnt_p.text = card["count"]
        cnt_p.alignment = pptx.enum.text.PP_ALIGN.CENTER
        cnt_p.font.name = "Aptos"
        cnt_p.font.size = Pt(22)
        cnt_p.font.bold = True
        cnt_p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)

        lbl_box = slide.shapes.add_textbox(Pt(bx + 8.0), Pt(by + 48.0), Pt(64.0), Pt(24.0))
        lbl_tf = lbl_box.text_frame
        lbl_p = lbl_tf.paragraphs[0]
        lbl_p.text = card["label"]
        lbl_p.alignment = pptx.enum.text.PP_ALIGN.CENTER
        lbl_p.font.name = "Aptos"
        lbl_p.font.size = Pt(9.5)
        lbl_p.font.bold = True
        lbl_p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)

        name_box = slide.shapes.add_textbox(Pt(bx + 80.0), Pt(by + 12.0), Pt(card_w - 90.0), Pt(24.0))
        n_tf = name_box.text_frame
        n_tf.word_wrap = True
        n_p = n_tf.paragraphs[0]
        n_p.text = card["name"]
        n_p.font.name = "Aptos"
        n_p.font.size = Pt(13)
        n_p.font.bold = True
        n_p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)

        # Authentic Library Line: Grey Side Fade R from Slide 6
        add_library_line(slide, lib_lines["side_fade_r"], bx + 80.0, by + 35.0, card_w - 95.0)

        reg_box = slide.shapes.add_textbox(Pt(bx + 80.0), Pt(by + 40.0), Pt(card_w - 90.0), Pt(38.0))
        r_tf = reg_box.text_frame
        r_tf.word_wrap = True
        r_p = r_tf.paragraphs[0]
        r_p.text = card["regions"]
        r_p.font.name = "Aptos"
        r_p.font.size = Pt(10)
        r_p.font.color.rgb = hex_to_rgb(ColorTokens.TEXT_SECONDARY)

    tot_shape = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(642.0), Pt(446.0), Pt(card_w), Pt(28.0))
    apply_solid_fill(tot_shape, ColorTokens.VIVID_INDIGO_SOFT)
    tot_shape.line.fill.background()
    t_tf = tot_shape.text_frame
    t_p = t_tf.paragraphs[0]
    t_p.text = s6_data["total_badge"]
    t_p.alignment = pptx.enum.text.PP_ALIGN.CENTER
    t_p.font.name = "Aptos"
    t_p.font.size = Pt(10.5)
    t_p.font.bold = True
    t_p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)


# ==============================================================================
# MAIN ENGINE EXECUTION
# ==============================================================================



def main():
    print(f"Loading input presentation: {INPUT_PATH}")
    with open(INPUT_PATH, "rb") as f:
        input_data = f.read()
    input_prs = pptx.Presentation(io.BytesIO(input_data))

    s1_data = extract_slide1_data(input_prs)
    s2_data = extract_slide2_data(input_prs)
    s3_data = extract_slide3_data(input_prs)
    s4_data = extract_slide4_data(input_prs)
    s5_data = extract_slide5_data(input_prs)
    s6_data = extract_slide6_data(input_prs)

    print(f"Extracted all 6 slides:")
    print(f"  Slide 1: {len(s1_data)} columns")
    print(f"  Slide 2: {len(s2_data['left_pairs'])} process rows + {s2_data['right_header']}")
    print(f"  Slide 3: {len(s3_data)} full-width process rows")
    print(f"  Slide 4: {s4_data['title'][:40]} | 5 rows + Deliverables + 4 KPIs")
    print(f"  Slide 5: {s5_data['title'][:40]} | 5 Workflows (A..E)")
    print(f"  Slide 6: {s6_data['title'][:40]} | 9 Persona cards + 2 badges")

    # Safely load Library Presentation
    lib_path_to_use = LIBRARY_PATH
    if not lib_path_to_use.exists() or os.access(lib_path_to_use, os.R_OK) is False:
        lib_path_to_use = LIBRARY_INSPECT_PATH
    try:
        with open(lib_path_to_use, "rb") as f:
            lib_bytes = f.read()
        lib_prs = pptx.Presentation(io.BytesIO(lib_bytes))
    except Exception:
        # Fallback to copy
        lib_prs = pptx.Presentation(LIBRARY_INSPECT_PATH)

    lib_s1 = lib_prs.slides[0]
    lib_s3 = lib_prs.slides[2]
    lib_s6 = lib_prs.slides[5]

    # Load 5 Authentic Library Lines from Slide 6
    lib_lines = {
        "side_fade_r": get_shape_by_name(lib_s6, "Grey Side Fade R"),
        "centre_fade_h": get_shape_by_name(lib_s6, "Grey Centre Fade H"),
        "solid_h": get_shape_by_name(lib_s6, "Grey Solid H"),
        "solid_v": get_shape_by_name(lib_s6, "Grey Solid V"),
        "centre_fade_v": get_shape_by_name(lib_s6, "Grey Centre Fade V"),
    }
    print(f"Loaded all 5 Library line prototypes from Slide 6: {list(lib_lines.keys())}")

    out_prs = pptx.Presentation()
    out_prs.slide_width = Pt(960)
    out_prs.slide_height = Pt(540)

    print("\n[Formatting Slide 1] 3-Column Floating Pill Cards with Library Grey Side Fade R lines...")
    build_slide1(out_prs, s1_data, lib_s1, lib_lines)

    print("[Formatting Slide 2] Split Layout with Library Grey Side Fade R lines...")
    build_slide2(out_prs, s2_data, lib_s1, lib_s3, lib_lines)

    print("[Formatting Slide 3] Full-Width Process Rows with Library Grey Side Fade R lines...")
    build_slide3(out_prs, s3_data, lib_s3, lib_lines)

    print("[Formatting Slide 4] Phase 1 Assessment with Library Grey Centre Fade H & Side Fade R lines...")
    build_slide4(out_prs, s4_data, lib_s1, lib_s3, lib_lines)

    print("[Formatting Slide 5] Phase 3 Five Multi-Agent Pilot Workflows...")
    build_slide5(out_prs, s5_data, lib_s3, lib_lines)

    print("[Formatting Slide 6] WS2 Team Structure: 25 FTEs (3x3 Grid)...")
    build_slide6(out_prs, s6_data, lib_s3, lib_lines)

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    buf = io.BytesIO()
    out_prs.save(buf)
    raw_bytes = buf.getvalue()

    print("[Enriching PPTX Package] Embedding authentic SVG vector icons and updating relationships...")
    file_bytes = inject_svg_assets(raw_bytes, AFTER_PATH)

    saved_paths = []
    for target in [OUTPUT_PATH, LOCAL_COPY_PATH]:
        try:
            with open(target, "wb") as f:
                f.write(file_bytes)
            saved_paths.append(str(target))
        except PermissionError:
            alt_target = target.with_name(target.stem + "_v2" + target.suffix)
            with open(alt_target, "wb") as f:
                f.write(file_bytes)
            saved_paths.append(str(alt_target))

    print(f"\nSuccessfully generated formatted presentations with Library line prototypes at:")
    for sp in saved_paths:
        print(f"  - {sp}")
    print(f"Total Slides: {len(out_prs.slides)}")


if __name__ == "__main__":
    main()

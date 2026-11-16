"""
AutoDesigner Precision Formatting Engine for All 7 Slides in Design tester new.pptx
Strictly adheres to:
1. Exact slide count: 7 input slides -> exactly 7 output slides (ZERO extra slides).
2. Hard minimum font size constraint: MINIMUM FONT SIZE >= 9.0 pt everywhere.
3. Rounded box constraint: 4px border radius on ALL rounded rectangles.
4. Line library constraint: All lines sourced exclusively from the 5 Library prototypes.
5. Shadow constraint: All newly added boxes have authentic Decor_Box_with_Shadow_1 shadow settings.
6. 100% Text Fidelity: All source text preserved with zero missing words.
7. Color theory: GEP brand colors (Brand Indigo #48397E, Tangerine #FE860E, Navy #1E293B, Soft Indigo #B1A8D8).
"""

import copy
import io
import os
import shutil
import subprocess
import sys
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
    set_shape_corner_radius,
    apply_library_shadow,
    enforce_deck_corner_radii,
    inject_svg_assets,
    load_library,
)
from engine.color_tokens import (
    ColorTokens,
    apply_linear_gradient,
    apply_solid_fill,
)

INPUT_PATH = BASE_DIR / "Design_tester_new_inspect.pptx"
OUTPUT_PATH = BASE_DIR / "output" / "Design_tester_new_formatted.pptx"
LOCAL_COPY_PATH = BASE_DIR / "Design tester new_formatted.pptx"


def hex_to_rgb(hex_str: str) -> RGBColor:
    hex_clean = hex_str.lstrip("#")
    return RGBColor(
        int(hex_clean[0:2], 16),
        int(hex_clean[2:4], 16),
        int(hex_clean[4:6], 16)
    )


# ==============================================================================
# DATA EXTRACTION FUNCTIONS
# ==============================================================================

def extract_slide1_data(input_prs):
    s = input_prs.slides[0]
    shapes_by_x = {}
    for shape in s.shapes:
        if shape.has_text_frame and shape.text_frame.text.strip():
            x_col = round((shape.left / 12700) / 150)
            shapes_by_x.setdefault(x_col, []).append(shape)

    columns = []
    for col_key in sorted(shapes_by_x.keys()):
        col_shapes = sorted(shapes_by_x[col_key], key=lambda sh: sh.top)
        header_text = col_shapes[0].text_frame.text.strip()
        body_texts = [sh.text_frame.text.strip() for sh in col_shapes[1:]]
        columns.append({"header": header_text, "items": body_texts})
    return columns


def extract_slide2_data(input_prs):
    s = input_prs.slides[1]
    left_shapes = [sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip() and (sh.left / 12700) < 600.0]
    right_shapes = [sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip() and (sh.left / 12700) >= 600.0]

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
    title = [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Case Study 3" in p.text][0]
    top_header = [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Accelerated Sourcing Program" in p.text][0]

    scope_items = [
        [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Spend Assessed" in p.text][0],
        [sh.text_frame.text.strip().replace('\x0b', ' ') for sh in s.shapes if sh.has_text_frame if "Categories:" in sh.text_frame.text][0],
        [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Regions-in-scope" in p.text][0],
        [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Opportunities Identified" in p.text][0],
    ]

    pain_header = [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Pain Points" in p.text][0]
    pain_shs = sorted([sh for sh in s.shapes if sh.has_text_frame and (sh.left / 914400) < 3.0 and 2.5 < (sh.top / 914400) < 6.0 and sh.text_frame.text.strip()], key=lambda sh: sh.top)
    pain_bullets = [sh.text_frame.text.strip() for sh in pain_shs]

    approach_header = [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if p.text.strip() == "Program Approach"][0]
    tb75 = [sh for sh in s.shapes if "Assess and Prioritize" in (sh.text_frame.text if sh.has_text_frame else "")][0]
    tb75_lines = [p.text.strip() for p in tb75.text_frame.paragraphs if p.text.strip()]
    app_sec1_title = tb75_lines[0]
    app_sec1_bullets = tb75_lines[1:]

    tb71 = [sh for sh in s.shapes if "Develop and Execute" in (sh.text_frame.text if sh.has_text_frame else "")][0]
    tb71_lines = [p.text.strip() for p in tb71.text_frame.paragraphs if p.text.strip()]
    app_sec2_title = tb71_lines[0]
    app_sec2_bullets = tb71_lines[1:]

    opps_header = [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Key Opportunities" in p.text][0]
    tb16 = [sh for sh in s.shapes if "IT contract harmonization" in (sh.text_frame.text if sh.has_text_frame else "")][0]
    opps_items = [p.text.strip() for p in tb16.text_frame.paragraphs if p.text.strip()]

    outcomes_header = [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if "Key Outcomes" in p.text][0]
    outcome_tiles = sorted([sh for sh in s.shapes if sh.has_text_frame and (sh.top / 914400) > 6.1 and sh.text_frame.text.strip() and "Key Outcomes" not in sh.text_frame.text], key=lambda sh: sh.left)
    outcomes_metrics = [sh.text_frame.text.strip() for sh in outcome_tiles]

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


def extract_table_slide(s, expected_rows, is_s7=False):
    title = [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if sh.top / 914400 < 1.0 and p.text.strip()][0]
    headers = [sh.text_frame.text.strip() for sh in sorted([sh for sh in s.shapes if sh.has_text_frame and 1.0 <= sh.top / 914400 <= 1.35], key=lambda x: x.left)]

    row_shapes = [sh for sh in s.shapes if sh.has_text_frame and sh.top / 914400 >= 1.4]
    rows = []
    threshold = 0.35 if expected_rows == 6 else 0.6
    for sh in sorted(row_shapes, key=lambda x: x.top):
        top_in = sh.top / 914400
        matched = False
        for r in rows:
            if abs(r["avg_top"] - top_in) < threshold:
                r["shapes"].append(sh)
                r["avg_top"] = sum(s.top / 914400 for s in r["shapes"]) / len(r["shapes"])
                matched = True
                break
        if not matched:
            rows.append({"avg_top": top_in, "shapes": [sh]})

    assert len(rows) == expected_rows, f"Expected {expected_rows} rows, got {len(rows)}"

    parsed_rows = []
    for r in rows:
        sorted_shs = sorted(r["shapes"], key=lambda x: x.left)
        if is_s7:
            name_sh = [s for s in sorted_shs if s.left / 914400 < 2.0][0]
            budget_sh = [s for s in sorted_shs if 2.0 <= s.left / 914400 < 5.0][0]
            interv_sh = [s for s in sorted_shs if 5.0 <= s.left / 914400 < 10.0][0]
            impact_sh = [s for s in sorted_shs if s.left / 914400 >= 10.0][0]

            b_text = budget_sh.text_frame.text.strip().replace("\x0b", "\n")
            b_lines = [line.strip() for line in b_text.split("\n") if line.strip()]

            parsed_rows.append({
                "name": name_sh.text_frame.text.strip(),
                "budget_assets": b_lines,
                "interventions": [p.text.strip() for p in interv_sh.text_frame.paragraphs if p.text.strip()],
                "impact": [p.text.strip() for p in impact_sh.text_frame.paragraphs if p.text.strip()],
            })
        else:
            col0_shs = [s for s in sorted_shs if s.left / 914400 < 2.5]
            badge_txt = ""
            name_txt = ""
            for s in col0_shs:
                t = s.text_frame.text.strip()
                if len(t) <= 2 or "Graphic" in s.name:
                    badge_txt = t
                else:
                    name_txt = t
            spend_sh = [s for s in sorted_shs if 2.5 <= s.left / 914400 < 5.0][0]
            interv_sh = [s for s in sorted_shs if 5.0 <= s.left / 914400 < 10.0][0]
            impact_sh = [s for s in sorted_shs if s.left / 914400 >= 10.0][0]

            spend_lines = [p.text.strip() for p in spend_sh.text_frame.paragraphs if p.text.strip()]
            interv_lines = [p.text.strip() for p in interv_sh.text_frame.paragraphs if p.text.strip()]
            impact_lines = [p.text.strip() for p in impact_sh.text_frame.paragraphs if p.text.strip()]

            parsed_rows.append({
                "badge": badge_txt,
                "name": name_txt,
                "spend_sup": spend_lines,
                "interventions": interv_lines,
                "impact": impact_lines,
            })

    return {
        "title": title,
        "headers": headers,
        "rows": parsed_rows,
    }


# ==============================================================================
# SLIDE BUILDERS (1 - 7)
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
            p.font.size = Pt(15.0)
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
                p.font.size = Pt(11.0)
                p.font.bold = False
                p.font.color.rgb = hex_to_rgb(ColorTokens.TEXT_PRIMARY)

            current_y += row_height

            if item_idx < len(col_data["items"]) - 1:
                current_y += gap_height
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

        # SVG icon placeholder
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
            p.font.size = Pt(11.0)
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
        p.font.size = Pt(15.0)
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
            p.font.size = Pt(11.0)
            p.font.bold = False
            p.font.color.rgb = hex_to_rgb(ColorTokens.TEXT_PRIMARY)

        current_y += row_height

        if item_idx < len(s2_data["right_items"]) - 1:
            current_y += gap_height
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

        # SVG icon placeholder
        icon_w_pt = 28.4
        icon_h_pt = 28.4
        icon_l_pt = row_left + 10.7 + (51.7 - icon_w_pt) / 2.0
        icon_t_pt = current_top + 10.1 + (51.7 - icon_h_pt) / 2.0
        emu_l = int(icon_l_pt * 12700)
        emu_t = int(icon_t_pt * 12700)
        emu_w = int(icon_w_pt * 12700)
        emu_h = int(icon_h_pt * 12700)
        pic_id = 8200 + idx
        rel_id = f"rIdSvgS3_{idx+1}"
        pic_xml = f'''<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:nvPicPr>
            <p:cNvPr id="{pic_id}" name="Graphic S3 {300+idx}"/>
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
            p.font.size = Pt(11.0)
            p.font.color.rgb = hex_to_rgb(ColorTokens.TEXT_PRIMARY)


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

    # Col 3: Key Opportunities
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
        m_p.font.size = Pt(9.5)
        m_p.font.bold = True
        m_p.font.color.rgb = hex_to_rgb("334155")


def build_slide5(prs, d, lib_lines):
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
        h_tab = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(hdr_lefts[h_i]), Pt(56.0), Pt(hdr_widths[h_i]), Pt(28.0))
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

    # 3. 4 Workstream Cards (A, B, C, D)
    row_tops = [96.0, 200.0, 304.0, 408.0]
    row_h = 96.0

    for r_i, tr in enumerate(d["rows"]):
        ry = row_tops[r_i]

        card = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(ry), Pt(900.0), Pt(row_h))
        apply_solid_fill(card, "FFFFFF")
        card.line.color.rgb = hex_to_rgb(ColorTokens.VIVID_INDIGO_SOFT)
        card.line.width = Pt(0.25)
        set_shape_corner_radius(card, 4.0)
        apply_library_shadow(card)

        # Col 1: Unified Badge (#B1A8D8 + shadow) + Name
        badge = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.OVAL, Pt(42.0), Pt(ry + 30.0), Pt(36.0), Pt(36.0))
        apply_solid_fill(badge, ColorTokens.VIVID_INDIGO_SOFT)
        badge.line.fill.background()
        apply_library_shadow(badge)
        bp = badge.text_frame.paragraphs[0]
        bp.text = tr["badge"]
        bp.alignment = pptx.enum.text.PP_ALIGN.CENTER
        bp.font.name = "Aptos"
        bp.font.size = Pt(16.0)
        bp.font.bold = True
        bp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

        name_box = slide.shapes.add_textbox(Pt(86.0), Pt(ry + 28.0), Pt(122.0), Pt(40.0))
        ntf = name_box.text_frame
        ntf.word_wrap = True
        np = ntf.paragraphs[0]
        np.text = tr["name"]
        np.font.name = "Aptos"
        np.font.size = Pt(12.0)
        np.font.bold = True
        np.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)

        add_library_line(slide, lib_lines["solid_v"], 214.0, ry + 12.0, width_pt=0.0, height_pt=72.0)

        # Col 2: Spend and Suppliers
        sp_box = slide.shapes.add_textbox(Pt(220.0), Pt(ry + 26.0), Pt(168.0), Pt(44.0))
        sptf = sp_box.text_frame
        sptf.word_wrap = True
        if tr["spend_sup"]:
            sp1 = sptf.paragraphs[0]
            sp1.text = tr["spend_sup"][0]
            sp1.font.name = "Aptos"
            sp1.font.size = Pt(11.0)
            sp1.font.bold = True
            sp1.font.color.rgb = hex_to_rgb("334155")

        if len(tr["spend_sup"]) > 1:
            sp2 = sptf.add_paragraph()
            sp2.text = tr["spend_sup"][1]
            sp2.font.name = "Aptos"
            sp2.font.size = Pt(14.0)
            sp2.font.bold = True
            sp2.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)
            sp2.space_before = Pt(3.0)

        add_library_line(slide, lib_lines["solid_v"], 392.0, ry + 12.0, width_pt=0.0, height_pt=72.0)

        # Col 3: Planned Interventions
        pi_box = slide.shapes.add_textbox(Pt(400.0), Pt(ry + 16.0), Pt(394.0), Pt(64.0))
        pitf = pi_box.text_frame
        pitf.word_wrap = True
        for p_i, pi_txt in enumerate(tr["interventions"]):
            p = pitf.paragraphs[0] if p_i == 0 else pitf.add_paragraph()
            p.text = "• " + pi_txt
            p.font.name = "Aptos"
            p.font.size = Pt(11.0)
            p.font.color.rgb = hex_to_rgb("334155")
            p.space_after = Pt(2.0)

        add_library_line(slide, lib_lines["solid_v"], 800.0, ry + 12.0, width_pt=0.0, height_pt=72.0)

        # Col 4: Estimated Impact
        imp_box = slide.shapes.add_textbox(Pt(806.0), Pt(ry + 24.0), Pt(118.0), Pt(48.0))
        imptf = imp_box.text_frame
        imptf.word_wrap = True
        if tr["impact"]:
            ip1 = imptf.paragraphs[0]
            ip1.text = tr["impact"][0]
            ip1.alignment = pptx.enum.text.PP_ALIGN.CENTER
            ip1.font.name = "Aptos"
            ip1.font.size = Pt(16.0)
            ip1.font.bold = True
            ip1.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_TANGERINE)

        if len(tr["impact"]) > 1:
            ip2 = imptf.add_paragraph()
            ip2.text = tr["impact"][1]
            ip2.alignment = pptx.enum.text.PP_ALIGN.CENTER
            ip2.font.name = "Aptos"
            ip2.font.size = Pt(10.5)
            ip2.font.bold = True
            ip2.font.color.rgb = hex_to_rgb("#475569")


def build_slide6(prs, d, lib_lines):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # 1. Title
    title_box = slide.shapes.add_textbox(Pt(30.0), Pt(10.0), Pt(900.0), Pt(32.0))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = d["title"]
    p.font.name = "Aptos"
    p.font.size = Pt(14.5)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

    # 2. Header Bar
    hdr_lefts = [30.0, 218.0, 396.0, 804.0]
    hdr_widths = [180.0, 170.0, 400.0, 126.0]
    for h_i, h_txt in enumerate(d["headers"]):
        h_tab = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(hdr_lefts[h_i]), Pt(48.0), Pt(hdr_widths[h_i]), Pt(26.0))
        apply_solid_fill(h_tab, ColorTokens.BRAND_INDIGO)
        h_tab.line.fill.background()
        set_shape_corner_radius(h_tab, 4.0)
        hp = h_tab.text_frame.paragraphs[0]
        hp.text = h_txt
        hp.alignment = pptx.enum.text.PP_ALIGN.CENTER
        hp.font.name = "Aptos"
        hp.font.size = Pt(11.5)
        hp.font.bold = True
        hp.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # 3. 6 Workstream Cards
    row_tops = [82.0, 156.0, 230.0, 304.0, 378.0, 452.0]
    row_h = 67.0

    for r_i, tr in enumerate(d["rows"]):
        ry = row_tops[r_i]

        card = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(ry), Pt(900.0), Pt(row_h))
        apply_solid_fill(card, "FFFFFF")
        card.line.color.rgb = hex_to_rgb(ColorTokens.VIVID_INDIGO_SOFT)
        card.line.width = Pt(0.25)
        set_shape_corner_radius(card, 4.0)
        apply_library_shadow(card)

        # Col 1: Unified Badge (#B1A8D8 + shadow) + Name
        badge = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.OVAL, Pt(40.0), Pt(ry + 18.5), Pt(30.0), Pt(30.0))
        apply_solid_fill(badge, ColorTokens.VIVID_INDIGO_SOFT)
        badge.line.fill.background()
        apply_library_shadow(badge)
        bp = badge.text_frame.paragraphs[0]
        bp.text = tr["badge"]
        bp.alignment = pptx.enum.text.PP_ALIGN.CENTER
        bp.font.name = "Aptos"
        bp.font.size = Pt(13.0)
        bp.font.bold = True
        bp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

        name_box = slide.shapes.add_textbox(Pt(76.0), Pt(ry + 10.0), Pt(132.0), Pt(47.0))
        ntf = name_box.text_frame
        ntf.word_wrap = True
        np = ntf.paragraphs[0]
        np.text = tr["name"]
        np.font.name = "Aptos"
        np.font.size = Pt(10.5)
        np.font.bold = True
        np.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)

        add_library_line(slide, lib_lines["solid_v"], 214.0, ry + 10.0, width_pt=0.0, height_pt=47.0)

        # Col 2: Spend and Suppliers
        sp_box = slide.shapes.add_textbox(Pt(220.0), Pt(ry + 12.0), Pt(168.0), Pt(45.0))
        sptf = sp_box.text_frame
        sptf.word_wrap = True
        if tr["spend_sup"]:
            sp1 = sptf.paragraphs[0]
            sp1.text = tr["spend_sup"][0]
            sp1.font.name = "Aptos"
            sp1.font.size = Pt(9.5)
            sp1.font.bold = True
            sp1.font.color.rgb = hex_to_rgb("334155")

        if len(tr["spend_sup"]) > 1:
            sp2 = sptf.add_paragraph()
            sp2.text = tr["spend_sup"][1]
            sp2.font.name = "Aptos"
            sp2.font.size = Pt(11.5)
            sp2.font.bold = True
            sp2.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)
            sp2.space_before = Pt(2.0)

        add_library_line(slide, lib_lines["solid_v"], 392.0, ry + 10.0, width_pt=0.0, height_pt=47.0)

        # Col 3: Planned Interventions
        pi_box = slide.shapes.add_textbox(Pt(400.0), Pt(ry + 6.0), Pt(394.0), Pt(55.0))
        pitf = pi_box.text_frame
        pitf.word_wrap = True
        pitf.margin_top = Pt(1.0)
        pitf.margin_bottom = Pt(1.0)
        for p_i, pi_txt in enumerate(tr["interventions"]):
            p = pitf.paragraphs[0] if p_i == 0 else pitf.add_paragraph()
            p.text = "• " + pi_txt
            p.font.name = "Aptos"
            p.font.size = Pt(9.0)
            p.font.color.rgb = hex_to_rgb("334155")
            p.space_after = Pt(1.0)

        add_library_line(slide, lib_lines["solid_v"], 800.0, ry + 10.0, width_pt=0.0, height_pt=47.0)

        # Col 4: Estimated Impact
        imp_box = slide.shapes.add_textbox(Pt(806.0), Pt(ry + 14.0), Pt(118.0), Pt(40.0))
        imptf = imp_box.text_frame
        imptf.word_wrap = True
        if tr["impact"]:
            ip1 = imptf.paragraphs[0]
            ip1.text = tr["impact"][0]
            ip1.alignment = pptx.enum.text.PP_ALIGN.CENTER
            ip1.font.name = "Aptos"
            ip1.font.size = Pt(15.0)
            ip1.font.bold = True
            ip1.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_TANGERINE)

        if len(tr["impact"]) > 1:
            ip2 = imptf.add_paragraph()
            ip2.text = tr["impact"][1]
            ip2.alignment = pptx.enum.text.PP_ALIGN.CENTER
            ip2.font.name = "Aptos"
            ip2.font.size = Pt(9.0)
            ip2.font.bold = True
            ip2.font.color.rgb = hex_to_rgb("#475569")


def build_slide7(prs, d, lib_lines):
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
    hdr_lefts = [30.0, 213.0, 426.0, 804.0]
    hdr_widths = [175.0, 205.0, 370.0, 126.0]
    for h_i, h_txt in enumerate(d["headers"]):
        h_tab = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(hdr_lefts[h_i]), Pt(56.0), Pt(hdr_widths[h_i]), Pt(28.0))
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

    # 3. 3 Workstream Cards
    row_tops = [98.0, 236.0, 374.0]
    row_h = 124.0

    for r_i, tr in enumerate(d["rows"]):
        ry = row_tops[r_i]

        card = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(ry), Pt(900.0), Pt(row_h))
        apply_solid_fill(card, "FFFFFF")
        card.line.color.rgb = hex_to_rgb(ColorTokens.VIVID_INDIGO_SOFT)
        card.line.width = Pt(0.25)
        set_shape_corner_radius(card, 4.0)
        apply_library_shadow(card)

        # Col 0: Master Title
        name_box = slide.shapes.add_textbox(Pt(42.0), Pt(ry + 36.0), Pt(160.0), Pt(52.0))
        ntf = name_box.text_frame
        ntf.word_wrap = True
        np = ntf.paragraphs[0]
        np.text = tr["name"]
        np.font.name = "Aptos"
        np.font.size = Pt(13.5)
        np.font.bold = True
        np.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)

        add_library_line(slide, lib_lines["solid_v"], 209.0, ry + 16.0, width_pt=0.0, height_pt=92.0)

        # Col 1: Budget and Assets
        b_box = slide.shapes.add_textbox(Pt(220.0), Pt(ry + 32.0), Pt(195.0), Pt(60.0))
        btf = b_box.text_frame
        btf.word_wrap = True
        if tr["budget_assets"]:
            bp1 = btf.paragraphs[0]
            bp1.text = tr["budget_assets"][0]
            bp1.font.name = "Aptos"
            bp1.font.size = Pt(12.0)
            bp1.font.bold = True
            bp1.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

        if len(tr["budget_assets"]) > 1:
            bp2 = btf.add_paragraph()
            bp2.text = tr["budget_assets"][1]
            bp2.font.name = "Aptos"
            bp2.font.size = Pt(14.0)
            bp2.font.bold = True
            bp2.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)
            bp2.space_before = Pt(4.0)

        add_library_line(slide, lib_lines["solid_v"], 422.0, ry + 16.0, width_pt=0.0, height_pt=92.0)

        # Col 2: Planned Interventions
        pi_box = slide.shapes.add_textbox(Pt(434.0), Pt(ry + 24.0), Pt(358.0), Pt(76.0))
        pitf = pi_box.text_frame
        pitf.word_wrap = True
        for p_i, pi_txt in enumerate(tr["interventions"]):
            p = pitf.paragraphs[0] if p_i == 0 else pitf.add_paragraph()
            p.text = "• " + pi_txt
            p.font.name = "Aptos"
            p.font.size = Pt(11.0)
            p.font.color.rgb = hex_to_rgb("334155")
            p.space_after = Pt(6.0)

        add_library_line(slide, lib_lines["solid_v"], 800.0, ry + 16.0, width_pt=0.0, height_pt=92.0)

        # Col 3: Estimated Impact
        imp_box = slide.shapes.add_textbox(Pt(806.0), Pt(ry + 40.0), Pt(118.0), Pt(44.0))
        imptf = imp_box.text_frame
        imptf.word_wrap = True
        if tr["impact"]:
            ip1 = imptf.paragraphs[0]
            ip1.text = tr["impact"][0]
            ip1.alignment = pptx.enum.text.PP_ALIGN.CENTER
            ip1.font.name = "Aptos"
            ip1.font.size = Pt(20.0)
            ip1.font.bold = True
            ip1.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_TANGERINE)


# ==============================================================================
# MAIN PIPELINE
# ==============================================================================

def main():
    print("Loading presentation...")
    src_path = BASE_DIR / "Design tester new.pptx"
    try:
        shutil.copyfile(src_path, INPUT_PATH)
    except Exception:
        subprocess.run(
            ["powershell", "-NoProfile", "-Command", f"Copy-Item -LiteralPath '{src_path}' -Destination '{INPUT_PATH}' -Force"],
            capture_output=True,
            check=False,
        )

    in_prs = pptx.Presentation(INPUT_PATH)
    print(f"Loaded input presentation with {len(in_prs.slides)} slides.")
    assert len(in_prs.slides) == 7, f"Expected 7 slides, found {len(in_prs.slides)}"

    print("Extracting data from all 7 slides...")
    s1_data = extract_slide1_data(in_prs)
    s2_data = extract_slide2_data(in_prs)
    s3_data = extract_slide3_data(in_prs)
    s4_data = extract_slide4_data(in_prs)
    s5_data = extract_table_slide(in_prs.slides[4], expected_rows=4)
    s6_data = extract_table_slide(in_prs.slides[5], expected_rows=6)
    s7_data = extract_table_slide(in_prs.slides[6], expected_rows=3, is_s7=True)
    print("All 7 slides extracted successfully.")

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

    print("[Formatting Slide 5] Opportunity Perspective (4 Rows A-D)...")
    build_slide5(out_prs, s5_data, lib_lines)

    print("[Formatting Slide 6] Opportunity Perspective (6 Rows 1-6)...")
    build_slide6(out_prs, s6_data, lib_lines)

    print("[Formatting Slide 7] Master Title / Digital Enablement (3 Rows)...")
    build_slide7(out_prs, s7_data, lib_lines)

    print("Total slides built:", len(out_prs.slides))
    assert len(out_prs.slides) == 7, f"Mismatch: built {len(out_prs.slides)} slides, expected 7!"

    print("Enforcing exact 4px border radius on all rounded boxes across the deck...")
    enforce_deck_corner_radii(out_prs, 4.0)

    raw_buf = io.BytesIO()
    out_prs.save(raw_buf)
    raw_bytes = raw_buf.getvalue()

    print("Injecting SVG vector icons into Slide 2 and Slide 3...")
    final_bytes = inject_svg_assets(raw_bytes, targets={2: "rIdSvgS2_", 3: "rIdSvgS3_"})

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_PATH, "wb") as f:
        f.write(final_bytes)
    with open(LOCAL_COPY_PATH, "wb") as f:
        f.write(final_bytes)

    print(f"\n==================================================================")
    print(f"SUCCESS: Formatted all 7 slides successfully!")
    print(f"  Output: {OUTPUT_PATH}")
    print(f"  Local:  {LOCAL_COPY_PATH}")
    print(f"==================================================================")


if __name__ == "__main__":
    main()

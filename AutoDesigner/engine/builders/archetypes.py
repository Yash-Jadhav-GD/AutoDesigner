"""
Slide builders ported from the per-deck format scripts.
Each builder takes plain, JSON-serialisable slide data (see engine/builders/registry.py
for the archetype each one implements) and appends one styled slide to `prs`.
"""

import copy

import pptx
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.oxml import parse_xml

from ..color_tokens import (
    ColorTokens,
    hex_to_rgb,
    apply_linear_gradient,
    apply_solid_fill,
)
from .common import (
    add_library_line,
    add_connector_line,
    clean_element,
    get_shape_by_name,
    set_cell_border,
)


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

        elem = clean_element(copy.deepcopy(d1_hdr_proto._element))
        slide.shapes._spTree.append(elem)
        hdr = slide.shapes[len(slide.shapes) - 1]
        hdr.left = Pt(bx + 30.5)
        hdr.top = Pt(card_top - 15.7)
        hdr.width = Pt(215.4)
        hdr.height = Pt(33.0)
        apply_linear_gradient(hdr, style["start"], style["end"], angle_dml=10800000, with_shadow=False)
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

    elem = clean_element(copy.deepcopy(d1_hdr_proto._element))
    slide.shapes._spTree.append(elem)
    hdr = slide.shapes[len(slide.shapes) - 1]
    hdr.left = Pt(right_x + (right_w - 215.4) / 2.0)
    hdr.top = Pt(card_top - 15.7)
    hdr.width = Pt(215.4)
    hdr.height = Pt(33.0)
    apply_linear_gradient(hdr, ColorTokens.GRADIENT_TANGERINE_START, ColorTokens.GRADIENT_TANGERINE_END, angle_dml=10800000, with_shadow=False)
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


def build_slide7(prs, s7_data, lib_lines):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # 1. Title
    title_box = slide.shapes.add_textbox(Pt(30.2), Pt(14.0), Pt(900.0), Pt(32.0))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = s7_data["title"]
    p.font.name = "Aptos"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

    # 2. Top Banner: PaaS 2.0 Program Management Team
    top_w = 420.0
    top_h = 28.0
    top_l = (960.0 - top_w) / 2.0
    top_t = 48.0

    banner = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(top_l), Pt(top_t), Pt(top_w), Pt(top_h))
    apply_solid_fill(banner, ColorTokens.BRAND_INDIGO)
    banner.line.fill.background()
    p_b = banner.text_frame.paragraphs[0]
    p_b.text = s7_data["top_banner"]
    p_b.alignment = pptx.enum.text.PP_ALIGN.CENTER
    p_b.font.name = "Aptos"
    p_b.font.size = Pt(11.5)
    p_b.font.bold = True
    p_b.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # Org Tree Top Connector Lines
    add_connector_line(slide, 480.0, top_t + top_h, 480.0, 90.0)
    add_connector_line(slide, 135.0, 90.0, 825.0, 90.0)

    col_x_centers = [135.0, 365.0, 595.0, 825.0]
    for cx in col_x_centers:
        add_connector_line(slide, cx, 90.0, cx, 104.0)

    # 3. 4 Leadership Cards at top = 104.0, height = 76.0, width = 210.0
    card_w = 210.0
    card_h = 76.0
    card_top = 104.0
    card_lefts = [30.0, 260.0, 490.0, 720.0]

    for idx, lead in enumerate(s7_data["leadership"]):
        cl = card_lefts[idx]
        c_shape = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(cl), Pt(card_top), Pt(card_w), Pt(card_h))
        if lead["type"] == "spend_value":
            apply_solid_fill(c_shape, "EEF2FF")
            c_shape.line.color.rgb = hex_to_rgb("8F9BC7")
            c_shape.line.width = Pt(1.5)
        else:
            apply_solid_fill(c_shape, "F8FAFC")
            c_shape.line.color.rgb = hex_to_rgb("CBD5E1")
            c_shape.line.width = Pt(1.0)

        # Title
        tbox = slide.shapes.add_textbox(Pt(cl + 6.0), Pt(card_top + 6.0), Pt(card_w - 12.0), Pt(28.0))
        tf = tbox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = lead["title"]
        p.alignment = pptx.enum.text.PP_ALIGN.CENTER
        p.font.name = "Aptos"
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)

        if lead.get("sub"):
            sub_p = tf.add_paragraph()
            sub_p.text = lead["sub"]
            sub_p.alignment = pptx.enum.text.PP_ALIGN.CENTER
            sub_p.font.name = "Aptos"
            sub_p.font.size = Pt(8.5)
            sub_p.font.color.rgb = hex_to_rgb(ColorTokens.TEXT_SECONDARY)

        # FTE Box
        fte_box = slide.shapes.add_textbox(Pt(cl), Pt(card_top + 40.0), Pt(card_w), Pt(28.0))
        fte_tf = fte_box.text_frame
        fte_p = fte_tf.paragraphs[0]
        fte_p.text = lead["ftes"]
        fte_p.alignment = pptx.enum.text.PP_ALIGN.CENTER
        fte_p.font.name = "Aptos"
        fte_p.font.size = Pt(16)
        fte_p.font.bold = True
        if lead["type"] == "spend_value":
            fte_p.font.color.rgb = hex_to_rgb(ColorTokens.VIVID_TANGERINE_MID)
        else:
            fte_p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)

    # 4. Connector Lines to Delivery Pools
    add_connector_line(slide, 135.0, card_top + card_h, 135.0, card_top + card_h + 10.0)
    add_connector_line(slide, 365.0, card_top + card_h, 365.0, card_top + card_h + 10.0)
    add_connector_line(slide, 135.0, card_top + card_h + 10.0, 365.0, card_top + card_h + 10.0)
    add_connector_line(slide, 250.0, card_top + card_h + 10.0, 250.0, 200.0)

    add_connector_line(slide, 595.0, card_top + card_h, 595.0, 200.0)
    add_connector_line(slide, 825.0, card_top + card_h, 825.0, 200.0)

    # 5. Delivery Pools Level (y = 200.0, height = 276.0)
    pool_y = 200.0
    pool_h = 276.0

    # Pool 1: Portfolio WL1 Delivery Pool
    p1 = s7_data["pool1"]
    p1_box = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(pool_y), Pt(440.0), Pt(pool_h))
    apply_solid_fill(p1_box, "FFFFFF")
    p1_box.line.color.rgb = hex_to_rgb("CBD5E1")

    p1_title_box = slide.shapes.add_textbox(Pt(44.0), Pt(pool_y + 8.0), Pt(260.0), Pt(22.0))
    p1_t_p = p1_title_box.text_frame.paragraphs[0]
    p1_t_p.text = p1["title"]
    p1_t_p.font.name = "Aptos"
    p1_t_p.font.size = Pt(12)
    p1_t_p.font.bold = True
    p1_t_p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)

    p1_badge = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(314.0), Pt(pool_y + 8.0), Pt(142.0), Pt(20.0))
    apply_solid_fill(p1_badge, ColorTokens.VIVID_INDIGO_SOFT)
    p1_badge.line.fill.background()
    p1_b_p = p1_badge.text_frame.paragraphs[0]
    p1_b_p.text = p1["badge"]
    p1_b_p.alignment = pptx.enum.text.PP_ALIGN.CENTER
    p1_b_p.font.name = "Aptos"
    p1_b_p.font.size = Pt(9.5)
    p1_b_p.font.bold = True
    p1_b_p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)

    p1_sub_box = slide.shapes.add_textbox(Pt(44.0), Pt(pool_y + 30.0), Pt(412.0), Pt(32.0))
    p1_sub_tf = p1_sub_box.text_frame
    p1_sub_tf.word_wrap = True
    p1_s_p = p1_sub_tf.paragraphs[0]
    p1_s_p.text = p1["sub"]
    p1_s_p.font.name = "Aptos"
    p1_s_p.font.size = Pt(8.5)
    p1_s_p.font.color.rgb = hex_to_rgb(ColorTokens.TEXT_SECONDARY)

    add_library_line(slide, lib_lines["side_fade_r"], 44.0, pool_y + 66.0, 412.0)

    # 6 Category Cards
    c_w = 130.0
    c_h = 86.0
    c_lefts = [44.0, 184.0, 324.0]
    c_tops = [pool_y + 78.0, pool_y + 172.0]
    for c_i, card_text in enumerate(p1["cards"]):
        r_i = c_i // 3
        col_i = c_i % 3
        sh = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(c_lefts[col_i]), Pt(c_tops[r_i]), Pt(c_w), Pt(c_h))
        apply_solid_fill(sh, "F8FAFC")
        sh.line.color.rgb = hex_to_rgb("E2E8F0")
        sh_p = sh.text_frame.paragraphs[0]
        sh_p.text = card_text
        sh_p.alignment = pptx.enum.text.PP_ALIGN.CENTER
        sh_p.font.name = "Aptos"
        sh_p.font.size = Pt(10)
        sh_p.font.bold = True
        sh_p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)

    # Pool 2: BU/BG WL1 Teams
    p2 = s7_data["pool2"]
    p2_box = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(490.0), Pt(pool_y), Pt(210.0), Pt(pool_h))
    apply_solid_fill(p2_box, "FFFFFF")
    p2_box.line.color.rgb = hex_to_rgb("CBD5E1")

    p2_title_box = slide.shapes.add_textbox(Pt(500.0), Pt(pool_y + 8.0), Pt(190.0), Pt(20.0))
    p2_t_p = p2_title_box.text_frame.paragraphs[0]
    p2_t_p.text = p2["title"]
    p2_t_p.font.name = "Aptos"
    p2_t_p.font.size = Pt(11)
    p2_t_p.font.bold = True
    p2_t_p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)

    p2_badge = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(500.0), Pt(pool_y + 28.0), Pt(120.0), Pt(18.0))
    apply_solid_fill(p2_badge, ColorTokens.VIVID_INDIGO_SOFT)
    p2_badge.line.fill.background()
    p2_b_p = p2_badge.text_frame.paragraphs[0]
    p2_b_p.text = p2["badge"]
    p2_b_p.alignment = pptx.enum.text.PP_ALIGN.CENTER
    p2_b_p.font.name = "Aptos"
    p2_b_p.font.size = Pt(8.5)
    p2_b_p.font.bold = True
    p2_b_p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)

    p2_sub_box = slide.shapes.add_textbox(Pt(500.0), Pt(pool_y + 48.0), Pt(190.0), Pt(32.0))
    p2_sub_tf = p2_sub_box.text_frame
    p2_sub_tf.word_wrap = True
    p2_s_p = p2_sub_tf.paragraphs[0]
    p2_s_p.text = p2["sub"]
    p2_s_p.font.name = "Aptos"
    p2_s_p.font.size = Pt(8)
    p2_s_p.font.color.rgb = hex_to_rgb(ColorTokens.TEXT_SECONDARY)

    add_library_line(slide, lib_lines["side_fade_r"], 500.0, pool_y + 84.0, 190.0)

    # 8 Regional Pills
    r_w = 90.0
    r_h = 36.0
    r_step_y = 44.0
    r_base_y = pool_y + 94.0
    for r_i, reg_text in enumerate(p2["regions"]):
        col_i = r_i % 2
        row_i = r_i // 2
        rx = 500.0 if col_i == 0 else 598.0
        ry = r_base_y + row_i * r_step_y
        r_sh = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(rx), Pt(ry), Pt(r_w), Pt(r_h))
        apply_solid_fill(r_sh, "F8FAFC")
        r_sh.line.color.rgb = hex_to_rgb("E2E8F0")
        r_p = r_sh.text_frame.paragraphs[0]
        r_p.text = reg_text
        r_p.alignment = pptx.enum.text.PP_ALIGN.CENTER
        r_p.font.name = "Aptos"
        r_p.font.size = Pt(9)
        r_p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

    # Pool 3: Functional WL1 Teams
    p3 = s7_data["pool3"]
    p3_box = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(720.0), Pt(pool_y), Pt(210.0), Pt(pool_h))
    apply_solid_fill(p3_box, "FFFFFF")
    p3_box.line.color.rgb = hex_to_rgb("CBD5E1")

    p3_title_box = slide.shapes.add_textbox(Pt(730.0), Pt(pool_y + 8.0), Pt(190.0), Pt(20.0))
    p3_t_p = p3_title_box.text_frame.paragraphs[0]
    p3_t_p.text = p3["title"]
    p3_t_p.font.name = "Aptos"
    p3_t_p.font.size = Pt(11)
    p3_t_p.font.bold = True
    p3_t_p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)

    p3_badge = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(730.0), Pt(pool_y + 28.0), Pt(120.0), Pt(18.0))
    apply_solid_fill(p3_badge, ColorTokens.VIVID_INDIGO_SOFT)
    p3_badge.line.fill.background()
    p3_b_p = p3_badge.text_frame.paragraphs[0]
    p3_b_p.text = p3["badge"]
    p3_b_p.alignment = pptx.enum.text.PP_ALIGN.CENTER
    p3_b_p.font.name = "Aptos"
    p3_b_p.font.size = Pt(8.5)
    p3_b_p.font.bold = True
    p3_b_p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)

    p3_sub_box = slide.shapes.add_textbox(Pt(730.0), Pt(pool_y + 48.0), Pt(190.0), Pt(32.0))
    p3_sub_tf = p3_sub_box.text_frame
    p3_sub_tf.word_wrap = True
    p3_s_p = p3_sub_tf.paragraphs[0]
    p3_s_p.text = p3["sub"]
    p3_s_p.font.name = "Aptos"
    p3_s_p.font.size = Pt(8)
    p3_s_p.font.color.rgb = hex_to_rgb(ColorTokens.TEXT_SECONDARY)

    add_library_line(slide, lib_lines["side_fade_r"], 730.0, pool_y + 84.0, 190.0)

    # 3 Functional Cards
    f_w = 186.0
    f_h = 50.0
    f_step_y = 60.0
    f_base_y = pool_y + 98.0
    for f_i, f_text in enumerate(p3["cards"]):
        fy = f_base_y + f_i * f_step_y
        f_sh = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(732.0), Pt(fy), Pt(f_w), Pt(f_h))
        apply_solid_fill(f_sh, "F8FAFC")
        f_sh.line.color.rgb = hex_to_rgb("E2E8F0")
        f_p = f_sh.text_frame.paragraphs[0]
        f_p.text = f_text
        f_p.alignment = pptx.enum.text.PP_ALIGN.CENTER
        f_p.font.name = "Aptos"
        f_p.font.size = Pt(10)
        f_p.font.bold = True
        f_p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)

    # 6. Bottom Legend
    leg = s7_data["legend"]
    sw1 = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(492.0), Pt(40.0), Pt(18.0))
    apply_solid_fill(sw1, "EEF2FF")
    sw1.line.color.rgb = hex_to_rgb("8F9BC7")

    l1_box = slide.shapes.add_textbox(Pt(76.0), Pt(490.0), Pt(180.0), Pt(22.0))
    l1_p = l1_box.text_frame.paragraphs[0]
    l1_p.text = leg["l1"]
    l1_p.font.name = "Aptos"
    l1_p.font.size = Pt(9.5)
    l1_p.font.bold = True
    l1_p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)

    sw2 = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(260.0), Pt(492.0), Pt(40.0), Pt(18.0))
    apply_solid_fill(sw2, "F8FAFC")
    sw2.line.color.rgb = hex_to_rgb("CBD5E1")

    l2_box = slide.shapes.add_textbox(Pt(306.0), Pt(490.0), Pt(180.0), Pt(22.0))
    l2_p = l2_box.text_frame.paragraphs[0]
    l2_p.text = leg["l2"]
    l2_p.font.name = "Aptos"
    l2_p.font.size = Pt(9.5)
    l2_p.font.bold = True
    l2_p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)


def build_slide8(prs, s8_data):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # 1. Title
    title_box = slide.shapes.add_textbox(Pt(30.2), Pt(6.0), Pt(900.0), Pt(38.0))
    tf = title_box.text_frame
    tf.word_wrap = True
    tf.margin_top = Pt(0.0)
    tf.margin_bottom = Pt(0.0)
    tf.margin_left = Pt(0.0)
    tf.margin_right = Pt(0.0)
    p = tf.paragraphs[0]
    p.text = s8_data["title"]
    p.font.name = "Aptos"
    p.font.size = Pt(15.0)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

    # 2. Callout Note Bar
    note_box = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.2), Pt(46.0), Pt(900.0), Pt(22.0))
    apply_solid_fill(note_box, "F1F5F9")
    note_box.line.color.rgb = hex_to_rgb("CBD5E1")
    n_tf = note_box.text_frame
    n_p = n_tf.paragraphs[0]
    n_p.text = s8_data["note"]
    n_p.alignment = pptx.enum.text.PP_ALIGN.CENTER
    n_p.font.name = "Aptos"
    n_p.font.size = Pt(9.0)
    n_p.font.color.rgb = hex_to_rgb("334155")

    # 3. Table (17 rows x 3 columns)
    t_data = s8_data["table_data"]
    num_rows = len(t_data)
    num_cols = len(t_data[0])
    table_shape = slide.shapes.add_table(num_rows, num_cols, Pt(30.2), Pt(72.0), Pt(900.0), Pt(388.0))
    tbl = table_shape.table
    tbl.columns[0].width = Pt(360.0)
    tbl.columns[1].width = Pt(180.0)
    tbl.columns[2].width = Pt(360.0)

    for r_idx, row_vals in enumerate(t_data):
        is_header = (r_idx == 0)
        is_total = (r_idx == num_rows - 1)
        is_packaging_sub = (7 <= r_idx <= 13)

        if is_header:
            bg_color = ColorTokens.BRAND_INDIGO
            text_color = RGBColor(0xFF, 0xFF, 0xFF)
            font_size = 9.5
            is_bold = True
        elif is_total:
            bg_color = ColorTokens.BRAND_NAVY
            text_color = RGBColor(0xFF, 0xFF, 0xFF)
            font_size = 10.0
            is_bold = True
        elif is_packaging_sub:
            bg_color = "FFFFFF"
            text_color = hex_to_rgb("475569")
            font_size = 8.5
            is_bold = False
        else:
            bg_color = "F8FAFC" if (r_idx % 2 == 1) else "FFFFFF"
            text_color = hex_to_rgb(ColorTokens.BRAND_ONYX)
            font_size = 9.0
            is_bold = True

        for c_idx, cell_paras in enumerate(row_vals):
            cell = tbl.cell(r_idx, c_idx)
            cell.fill.solid()
            r = int(bg_color[0:2], 16)
            g = int(bg_color[2:4], 16)
            b = int(bg_color[4:6], 16)
            cell.fill.fore_color.rgb = RGBColor(r, g, b)

            set_cell_border(cell, color="E2E8F0", width="6350")

            tf = cell.text_frame
            tf.word_wrap = True
            tf.margin_left = Pt(6.0) if c_idx == 0 else Pt(4.0)
            tf.margin_right = Pt(6.0)
            tf.margin_top = Pt(1.5)
            tf.margin_bottom = Pt(1.5)

            for p_idx, p_text in enumerate(cell_paras):
                p = tf.paragraphs[0] if p_idx == 0 else tf.add_paragraph()
                p.text = p_text
                p.font.name = "Aptos"
                p.font.size = Pt(font_size if not (c_idx == 2 and p_idx > 0) else 8.0)
                p.font.bold = is_bold if p_idx == 0 else False
                p.font.color.rgb = text_color if not (c_idx == 2 and p_idx > 0) else hex_to_rgb("64748B")

                if is_header or is_total:
                    p.alignment = pptx.enum.text.PP_ALIGN.CENTER if c_idx > 0 else pptx.enum.text.PP_ALIGN.LEFT
                elif c_idx == 1:
                    p.alignment = pptx.enum.text.PP_ALIGN.CENTER
                elif c_idx == 2:
                    p.alignment = pptx.enum.text.PP_ALIGN.CENTER if not (len(cell_paras) > 1) else pptx.enum.text.PP_ALIGN.LEFT
                else:
                    p.alignment = pptx.enum.text.PP_ALIGN.LEFT

    # 4. Footer Note
    f_box = slide.shapes.add_textbox(Pt(30.2), Pt(506.0), Pt(900.0), Pt(20.0))
    f_tf = f_box.text_frame
    f_p = f_tf.paragraphs[0]
    f_p.text = s8_data["footer"]
    f_p.font.name = "Aptos"
    f_p.font.size = Pt(8.5)
    f_p.font.italic = True
    f_p.font.color.rgb = hex_to_rgb("64748B")


def build_slide9(prs, s9_data):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # 1. Title
    title_box = slide.shapes.add_textbox(Pt(30.2), Pt(14.0), Pt(900.0), Pt(36.0))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = s9_data["title"]
    p.font.name = "Aptos"
    p.font.size = Pt(17.5)
    p.font.bold = True
    p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)

    # 2. Table (13 rows x 3 columns)
    t_data = s9_data["table_data"]
    num_rows = len(t_data)
    num_cols = len(t_data[0])
    table_shape = slide.shapes.add_table(num_rows, num_cols, Pt(30.2), Pt(58.0), Pt(900.0), Pt(454.0))
    tbl = table_shape.table
    tbl.columns[0].width = Pt(240.0)
    tbl.columns[1].width = Pt(110.0)
    tbl.columns[2].width = Pt(550.0)

    for r_idx, row_vals in enumerate(t_data):
        is_header = (r_idx == 0)
        is_total = (r_idx == num_rows - 1)
        is_tier1 = (r_idx == 1)
        is_bu_header = (r_idx == 2)
        is_bu_sub = (3 <= r_idx <= 9)
        is_poc = (r_idx == 10)
        is_si = (r_idx == 11)

        if is_header:
            bg_color = ColorTokens.BRAND_INDIGO
            text_color = RGBColor(0xFF, 0xFF, 0xFF)
            font_size = 11.0
            is_bold = True
        elif is_total:
            bg_color = ColorTokens.BRAND_NAVY
            text_color = RGBColor(0xFF, 0xFF, 0xFF)
            font_size = 11.5
            is_bold = True
        elif is_tier1 or is_poc:
            bg_color = "F1F0F8"
            text_color = hex_to_rgb(ColorTokens.BRAND_INDIGO)
            font_size = 10.0
            is_bold = True
        elif is_bu_header:
            bg_color = "E8E6F4"
            text_color = hex_to_rgb(ColorTokens.BRAND_INDIGO)
            font_size = 10.5
            is_bold = True
        elif is_bu_sub:
            bg_color = "FFFFFF" if (r_idx % 2 == 1) else "F8FAFC"
            text_color = hex_to_rgb("475569")
            font_size = 9.5
            is_bold = False
        else:
            bg_color = "FFFFFF"
            text_color = hex_to_rgb(ColorTokens.BRAND_INDIGO)
            font_size = 10.0
            is_bold = True

        for c_idx, cell_paras in enumerate(row_vals):
            cell = tbl.cell(r_idx, c_idx)
            cell.fill.solid()
            r = int(bg_color[0:2], 16)
            g = int(bg_color[2:4], 16)
            b = int(bg_color[4:6], 16)
            cell.fill.fore_color.rgb = RGBColor(r, g, b)

            set_cell_border(cell, color="E2E8F0", width="6350")

            tf = cell.text_frame
            tf.word_wrap = True
            tf.margin_left = Pt(10.0) if c_idx != 1 else Pt(4.0)
            tf.margin_right = Pt(10.0)
            tf.margin_top = Pt(4.0)
            tf.margin_bottom = Pt(4.0)

            for p_idx, p_text in enumerate(cell_paras):
                p = tf.paragraphs[0] if p_idx == 0 else tf.add_paragraph()
                p.text = p_text
                p.font.name = "Aptos"
                p.font.size = Pt(font_size if not (c_idx == 2 and p_idx > 0) else 9.0)
                p.font.bold = is_bold if (p_idx == 0 and c_idx < 2) else False
                p.font.color.rgb = text_color if not (c_idx == 2 and p_idx > 0) else hex_to_rgb("64748B")

                if is_header or is_total:
                    p.alignment = pptx.enum.text.PP_ALIGN.CENTER if c_idx == 1 else pptx.enum.text.PP_ALIGN.LEFT
                elif c_idx == 1:
                    p.alignment = pptx.enum.text.PP_ALIGN.CENTER
                else:
                    p.alignment = pptx.enum.text.PP_ALIGN.LEFT

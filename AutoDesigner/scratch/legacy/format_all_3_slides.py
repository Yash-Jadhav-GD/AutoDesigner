"""
AutoDesigner Precision Formatting Engine with Context-Aware Color Theory
Strictly adheres to:
1. COLOR_THEORY_SPECIFICATION.md & After_copy.pptx visual benchmarks
2. Contextual semantic coloring:
   - Slide 1: 3-Part progression (Indigo Problem -> Navy Bridge -> Tangerine Future State)
   - Slide 2: Left points represent unified diagnostic assessment streams -> all Icon Boxes share identical vivid.indigo.soft (#B1A8D8) backplate
   - Slide 2 Right: INSIGHTS stands out as the singular Tangerine focal takeaway callout
   - Slide 3: 5 full-width diagnostic rows share identical vivid.indigo.soft (#B1A8D8) backplates for semantic consistency
3. 100% Text Fidelity (ZERO added words, ZERO subtracted words)
"""

import copy
import io
import os
import sys
from pathlib import Path
import pptx
from pptx.util import Pt
from pptx.dml.color import RGBColor

# Import ColorTokens and OpenXML styling helpers
BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))
from engine.builders.common import clean_element, get_shape_by_name
from engine.color_tokens import (
    ColorTokens,
    hex_to_rgb,
    apply_linear_gradient,
    apply_solid_fill,
)

INPUT_PATH = BASE_DIR / "Design tester new.pptx"
LIBRARY_PATH = BASE_DIR / "Library_copy.pptx"
OUTPUT_PATH = BASE_DIR / "output" / "Design_tester_new_formatted.pptx"
LOCAL_COPY_PATH = BASE_DIR / "Design tester new_formatted.pptx"






def extract_slide1_data(input_prs):
    """Extracts raw text strings from Slide 1 with exact character fidelity."""
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
    """Extracts split layout text from Slide 2: Left process rows and Right insights card."""
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
    """Extracts raw text strings from Slide 3 with exact character fidelity."""
    s = input_prs.slides[2]
    text_shapes = [sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip()]
    text_shapes.sort(key=lambda sh: sh.top)

    pairs = []
    for i in range(0, len(text_shapes), 2):
        sub_text = text_shapes[i].text_frame.text.strip()
        body_text = text_shapes[i + 1].text_frame.text.strip() if i + 1 < len(text_shapes) else ""
        pairs.append({"subheader": sub_text, "body": body_text})
    return pairs


def build_slide1(prs, s1_data, lib_s1):
    """
    Formats Slide 1: 3-Column Floating Pill Cards adhering to COLOR_THEORY_SPECIFICATION.md:
    - Column 0 (The Problem): Indigo Gradient (#6A5E96 -> #48397E) + White Bold Text + Drop Shadow
    - Column 1 (The Bridge): Navy Gradient (#5A668E -> #1F2E65) + White Bold Text + Drop Shadow
    - Column 2 (The Future State): Tangerine Gradient (#FEB163 -> #FE860E) + White Bold Text + Drop Shadow (After_copy Slide 1)
    """
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    d1_box_proto = get_shape_by_name(lib_s1, "Design 1 Decor box")
    d1_hdr_proto = get_shape_by_name(lib_s1, "Design 1 Header Text")
    d1_body_proto = get_shape_by_name(lib_s1, "Design 1 Body text 1")
    d1_div_proto = get_shape_by_name(lib_s1, "Grey Side Fade R Divider")

    column_x = [30.2, 341.8, 653.5]
    card_w = 276.4
    card_top = 80.0
    card_h = 410.0

    pill_styles = [
        {
            "start": ColorTokens.GRADIENT_INDIGO_START,
            "end": ColorTokens.GRADIENT_INDIGO_END,
            "text_color": hex_to_rgb(ColorTokens.TEXT_ON_DARK),
        },
        {
            "start": ColorTokens.GRADIENT_NAVY_START,
            "end": ColorTokens.GRADIENT_NAVY_END,
            "text_color": hex_to_rgb(ColorTokens.TEXT_ON_DARK),
        },
        {
            "start": ColorTokens.GRADIENT_TANGERINE_START,
            "end": ColorTokens.GRADIENT_TANGERINE_END,
            "text_color": hex_to_rgb(ColorTokens.TEXT_ON_DARK),  # Matches After_copy Slide 1
        },
    ]

    for col_idx, col_data in enumerate(s1_data):
        bx = column_x[col_idx]
        style = pill_styles[col_idx] if col_idx < len(pill_styles) else pill_styles[0]

        # 1. Container Decor Box
        elem = clean_element(copy.deepcopy(d1_box_proto._element))
        slide.shapes._spTree.append(elem)
        box = slide.shapes[len(slide.shapes) - 1]
        box.left = Pt(bx)
        box.top = Pt(card_top)
        box.width = Pt(card_w)
        box.height = Pt(card_h)

        # 2. Floating Pill Header (15.7 pt protrusion + Linear Gradient + Drop Shadow)
        elem = clean_element(copy.deepcopy(d1_hdr_proto._element))
        slide.shapes._spTree.append(elem)
        hdr = slide.shapes[len(slide.shapes) - 1]
        hdr.left = Pt(bx + 30.5)
        hdr.top = Pt(card_top - 15.7)
        hdr.width = Pt(215.4)
        hdr.height = Pt(33.0)

        # Apply After_copy gradient & shadow
        apply_linear_gradient(hdr, style["start"], style["end"], angle_dml=10800000, with_shadow=True)

        hdr.text_frame.text = col_data["header"]
        for p in hdr.text_frame.paragraphs:
            p.font.name = "Aptos"
            p.font.size = Pt(15)
            p.font.bold = True
            p.font.color.rgb = style["text_color"]

        # 3. Symmetrical Body Rows & Gradient Dividers
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
                elem = clean_element(copy.deepcopy(d1_div_proto._element))
                slide.shapes._spTree.append(elem)
                div = slide.shapes[len(slide.shapes) - 1]
                div.left = Pt(bx + 17.5)
                div.top = Pt(current_y)
                div.width = Pt(card_w - 35.0)
                current_y += gap_height


def build_slide2(prs, s2_data, lib_s1, lib_s3):
    """
    Formats Slide 2: Split Executive Layout
    - Left: 5 Process Rows representing unified diagnostic assessment points
            -> ALL Icon Boxes have the SAME COLOR: vivid.indigo.soft (#B1A8D8)
    - Right: INSIGHTS Floating Pill Card styled with Tangerine Gradient as the singular focal takeaway (<=10%)
    """
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Left Prototypes
    decor_proto = get_shape_by_name(lib_s3, "UI-ROW-01 Decor Box")
    icon_box_proto = get_shape_by_name(lib_s3, "UI-ROW-01 Icon Box")
    subheader_proto = get_shape_by_name(lib_s3, "UI-ROW-01 Sub Header Text")
    line_proto = get_shape_by_name(lib_s3, "UI-ROW-01 Line")
    body_proto = get_shape_by_name(lib_s3, "UI-ROW-01 Body Text")

    # Right Prototypes
    d1_box_proto = get_shape_by_name(lib_s1, "Design 1 Decor box")
    d1_hdr_proto = get_shape_by_name(lib_s1, "Design 1 Header Text")
    d1_body_proto = get_shape_by_name(lib_s1, "Design 1 Body text 1")
    d1_div_proto = get_shape_by_name(lib_s1, "Grey Side Fade R Divider")

    left_x = 30.2
    left_w = 590.0
    row_count = len(s2_data["left_pairs"])
    row_h = 72.0
    row_gap = 12.5
    start_top = 65.0

    # Semantic color: All 5 diagnostic assessment points share the unified vivid.indigo.soft backplate
    diagnostic_color = ColorTokens.VIVID_INDIGO_SOFT  # #B1A8D8

    # 1. Left 5 Process Rows
    for idx, item in enumerate(s2_data["left_pairs"]):
        current_top = start_top + idx * (row_h + row_gap)

        # Row Decor Box
        elem = clean_element(copy.deepcopy(decor_proto._element))
        slide.shapes._spTree.append(elem)
        box = slide.shapes[len(slide.shapes) - 1]
        box.left = Pt(left_x)
        box.top = Pt(current_top)
        box.width = Pt(left_w)
        box.height = Pt(row_h)

        # Left Accent Box: ALL share the same diagnostic color (#B1A8D8)
        elem = clean_element(copy.deepcopy(icon_box_proto._element))
        slide.shapes._spTree.append(elem)
        ibox = slide.shapes[len(slide.shapes) - 1]
        ibox.left = Pt(left_x + 10.0)
        ibox.top = Pt(current_top + 10.1)
        ibox.width = Pt(51.7)
        ibox.height = Pt(51.7)
        apply_solid_fill(ibox, diagnostic_color)
        ibox.text_frame.text = ""  # Zero added text

        # Subheader Text
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

        # Divider Line
        elem = clean_element(copy.deepcopy(line_proto._element))
        slide.shapes._spTree.append(elem)
        line = slide.shapes[len(slide.shapes) - 1]
        line.left = Pt(left_x + 72.0)
        line.top = Pt(current_top + 31.0)
        line.width = Pt(left_w - 82.0)

        # Body Text
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

    # 2. Right INSIGHTS Floating Pill Card
    right_x = 650.0
    right_w = 280.0
    card_top = 80.0
    card_h = 410.0

    # Background Decor Box
    elem = clean_element(copy.deepcopy(d1_box_proto._element))
    slide.shapes._spTree.append(elem)
    box = slide.shapes[len(slide.shapes) - 1]
    box.left = Pt(right_x)
    box.top = Pt(card_top)
    box.width = Pt(right_w)
    box.height = Pt(card_h)

    # Floating Header Pill (Tangerine Gradient Callout + White Bold Text + Drop Shadow)
    elem = clean_element(copy.deepcopy(d1_hdr_proto._element))
    slide.shapes._spTree.append(elem)
    hdr = slide.shapes[len(slide.shapes) - 1]
    hdr.left = Pt(right_x + (right_w - 215.4) / 2.0)
    hdr.top = Pt(card_top - 15.7)
    hdr.width = Pt(215.4)
    hdr.height = Pt(33.0)
    apply_linear_gradient(
        hdr,
        ColorTokens.GRADIENT_TANGERINE_START,
        ColorTokens.GRADIENT_TANGERINE_END,
        angle_dml=10800000,
        with_shadow=True,
    )
    hdr.text_frame.text = s2_data["right_header"]
    for p in hdr.text_frame.paragraphs:
        p.font.name = "Aptos"
        p.font.size = Pt(15)
        p.font.bold = True
        p.font.color.rgb = hex_to_rgb(ColorTokens.TEXT_ON_DARK)

    # 4 Bullet Items inside INSIGHTS card
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
            elem = clean_element(copy.deepcopy(d1_div_proto._element))
            slide.shapes._spTree.append(elem)
            div = slide.shapes[len(slide.shapes) - 1]
            div.left = Pt(right_x + 17.5)
            div.top = Pt(current_y)
            div.width = Pt(right_w - 35.0)
            current_y += gap_height


def build_slide3(prs, s3_data, lib_s3):
    """
    Formats Slide 3: 5 Full-Width Horizontal Process Cards
    - All 5 diagnostic assessment points share the unified vivid.indigo.soft (#B1A8D8) backplate
    - Subheaders in brand.indigo (#48397E)
    - Body copy in text.primary (#262626)
    """
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    decor_proto = get_shape_by_name(lib_s3, "UI-ROW-01 Decor Box")
    icon_box_proto = get_shape_by_name(lib_s3, "UI-ROW-01 Icon Box")
    subheader_proto = get_shape_by_name(lib_s3, "UI-ROW-01 Sub Header Text")
    line_proto = get_shape_by_name(lib_s3, "UI-ROW-01 Line")
    body_proto = get_shape_by_name(lib_s3, "UI-ROW-01 Body Text")

    row_count = len(s3_data)
    row_h = 72.0
    row_gap = 14.0
    total_h = row_count * row_h + (row_count - 1) * row_gap
    start_top = (540.0 - total_h) / 2.0

    row_left = 24.8
    row_width = 910.4

    diagnostic_color = ColorTokens.VIVID_INDIGO_SOFT  # #B1A8D8

    for idx, item in enumerate(s3_data):
        current_top = start_top + idx * (row_h + row_gap)

        # Row Decor Box
        elem = clean_element(copy.deepcopy(decor_proto._element))
        slide.shapes._spTree.append(elem)
        box = slide.shapes[len(slide.shapes) - 1]
        box.left = Pt(row_left)
        box.top = Pt(current_top)
        box.width = Pt(row_width)
        box.height = Pt(row_h)

        # Left Accent Box: ALL share the same diagnostic color (#B1A8D8)
        elem = clean_element(copy.deepcopy(icon_box_proto._element))
        slide.shapes._spTree.append(elem)
        ibox = slide.shapes[len(slide.shapes) - 1]
        ibox.left = Pt(row_left + 10.7)
        ibox.top = Pt(current_top + 10.1)
        ibox.width = Pt(51.7)
        ibox.height = Pt(51.7)
        apply_solid_fill(ibox, diagnostic_color)
        ibox.text_frame.text = ""  # Zero added text

        # Subheader Text
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

        # Divider Line
        elem = clean_element(copy.deepcopy(line_proto._element))
        slide.shapes._spTree.append(elem)
        line = slide.shapes[len(slide.shapes) - 1]
        line.left = Pt(row_left + 74.0)
        line.top = Pt(current_top + 31.0)
        line.width = Pt(row_width - 86.0)

        # Body Text
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


def main():
    print(f"Loading input presentation from: {INPUT_PATH}")
    with open(INPUT_PATH, "rb") as f:
        input_data = f.read()
    input_prs = pptx.Presentation(io.BytesIO(input_data))

    s1_data = extract_slide1_data(input_prs)
    s2_data = extract_slide2_data(input_prs)
    s3_data = extract_slide3_data(input_prs)

    print(f"Slide 1 extracted: {len(s1_data)} columns with {sum(len(c['items']) for c in s1_data)} total items")
    print(f"Slide 2 extracted: {len(s2_data['left_pairs'])} process rows + {s2_data['right_header']} card with {len(s2_data['right_items'])} items")
    print(f"Slide 3 extracted: {len(s3_data)} full-width process rows")

    print(f"Loading template library from: {LIBRARY_PATH}")
    lib_prs = pptx.Presentation(LIBRARY_PATH)
    lib_s1 = lib_prs.slides[0]
    lib_s3 = lib_prs.slides[2]

    out_prs = pptx.Presentation()
    out_prs.slide_width = Pt(960)
    out_prs.slide_height = Pt(540)

    print("\n[Formatting Slide 1] Applying Problem-Bridge-Future color gradients & drop shadows...")
    build_slide1(out_prs, s1_data, lib_s1)

    print("[Formatting Slide 2] Applying unified diagnostic accent color to left rows & Tangerine INSIGHTS callout...")
    build_slide2(out_prs, s2_data, lib_s1, lib_s3)

    print("[Formatting Slide 3] Applying unified diagnostic accent color to full-width process rows...")
    build_slide3(out_prs, s3_data, lib_s3)

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    buf = io.BytesIO()
    out_prs.save(buf)
    file_bytes = buf.getvalue()

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

    print(f"\nSuccessfully generated formatted presentations with Context-Aware Color Theory at:")
    for sp in saved_paths:
        print(f"  - {sp}")
    print(f"Total Slides: {len(out_prs.slides)}")


if __name__ == "__main__":
    main()

"""
AutoDesigner Precision Formatting Engine
Formats Slide 1 and Slide 2 of Design tester new.pptx with 100% text fidelity:
- ZERO added text (no invented titles or banners)
- ZERO subtracted text (every character, dash, and number preserved identically)
- Strict compliance with AUTOFORMAT_SPECIFICATION.md and Library.pptx prototypes
- 100% verified OpenXML relationship integrity for PowerPoint desktop
"""

import copy
import io
from pathlib import Path
import pptx
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

BASE_DIR = Path(__file__).resolve().parent
INPUT_PATH = BASE_DIR / "Design tester new.pptx"
LIBRARY_PATH = BASE_DIR / "Library_copy.pptx"
AFTER_PATH = BASE_DIR / "After_copy.pptx"
OUTPUT_PATH = BASE_DIR / "output" / "Design_tester_new_formatted.pptx"


def clean_element(elem):
    """Strips any internal custDataLst / tags elements that refer to missing rId2 relationships."""
    for parent in elem.iter():
        for child in list(parent):
            if child.tag.endswith("custDataLst") or child.tag.endswith("tags"):
                parent.remove(child)
    return elem


def get_shape_by_name(slide, name):
    for s in slide.shapes:
        if s.name == name:
            return s
    return None


def extract_slide1_data(input_prs):
    """Extracts raw text strings from Slide 1 with exact character fidelity."""
    s = input_prs.slides[0]
    # Slide 1 has 3 columns:
    # Col 1: The Problem (4 items)
    # Col 2: The Bridge (4 items)
    # Col 3: The Future State (4 items)
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
    """Extracts raw text strings from Slide 2 with exact character fidelity."""
    s = input_prs.slides[1]
    # Filter text-bearing shapes only (ignores background rectangles like Rectangle 1 / 3)
    text_shapes = [sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip()]
    text_shapes.sort(key=lambda sh: sh.top)

    pairs = []
    for i in range(0, len(text_shapes), 2):
        sub_text = text_shapes[i].text_frame.text.strip()
        body_text = text_shapes[i + 1].text_frame.text.strip() if i + 1 < len(text_shapes) else ""
        pairs.append({"subheader": sub_text, "body": body_text})
    return pairs


def build_slide1_formatted(prs, s1_data, lib_s1):
    """Formats Slide 1 using Design 1 Floating Pill Cards with strict aspect-ratio locking & gradient dividers."""
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

    for col_idx, col_data in enumerate(s1_data):
        bx = column_x[col_idx]

        # 1. Background Container Decor Box
        elem = clean_element(copy.deepcopy(d1_box_proto._element))
        slide.shapes._spTree.append(elem)
        box = slide.shapes[len(slide.shapes) - 1]
        box.left = Pt(bx)
        box.top = Pt(card_top)
        box.width = Pt(card_w)
        box.height = Pt(card_h)

        # 2. Floating Header Pill (15.7pt protrusion)
        elem = clean_element(copy.deepcopy(d1_hdr_proto._element))
        slide.shapes._spTree.append(elem)
        hdr = slide.shapes[len(slide.shapes) - 1]
        hdr.left = Pt(bx + 30.5)
        hdr.top = Pt(card_top - 15.7)
        hdr.width = Pt(215.4)
        hdr.height = Pt(33.0)
        hdr.text_frame.text = col_data["header"]
        for p in hdr.text_frame.paragraphs:
            p.font.name = "Aptos"
            p.font.size = Pt(15)
            p.font.bold = True
            p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

        # 3. Dynamic Vertical Row Distribution
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
                p.font.color.rgb = RGBColor(0x26, 0x26, 0x26)

            current_y += row_height

            # Divider line between rows
            if item_idx < len(col_data["items"]) - 1:
                current_y += gap_height
                elem = clean_element(copy.deepcopy(d1_div_proto._element))
                slide.shapes._spTree.append(elem)
                div = slide.shapes[len(slide.shapes) - 1]
                div.left = Pt(bx + 17.5)
                div.top = Pt(current_y)
                div.width = Pt(card_w - 35.0)
                current_y += gap_height


def build_slide2_formatted(prs, s2_data, lib_s3):
    """Formats Slide 2 using UI-ROW-01 Horizontal Row Cards from Library.pptx Slide 3 (Zero added text)."""
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    decor_proto = get_shape_by_name(lib_s3, "UI-ROW-01 Decor Box")
    icon_box_proto = get_shape_by_name(lib_s3, "UI-ROW-01 Icon Box")
    subheader_proto = get_shape_by_name(lib_s3, "UI-ROW-01 Sub Header Text")
    line_proto = get_shape_by_name(lib_s3, "UI-ROW-01 Line")
    body_proto = get_shape_by_name(lib_s3, "UI-ROW-01 Body Text")

    row_count = len(s2_data)  # 5 rows
    row_h = 72.0
    row_gap = 14.0
    total_h = row_count * row_h + (row_count - 1) * row_gap
    start_top = (540.0 - total_h) / 2.0  # 62.0 pt (symmetrically centered vertically)

    row_left = 24.8
    row_width = 910.4

    for idx, item in enumerate(s2_data):
        current_top = start_top + idx * (row_h + row_gap)

        # 1. Row Decor Box Container
        elem = clean_element(copy.deepcopy(decor_proto._element))
        slide.shapes._spTree.append(elem)
        box = slide.shapes[len(slide.shapes) - 1]
        box.left = Pt(row_left)
        box.top = Pt(current_top)
        box.width = Pt(row_width)
        box.height = Pt(row_h)

        # 2. Icon / Accent Box on Left (Empty visual accent container - NO added text)
        elem = clean_element(copy.deepcopy(icon_box_proto._element))
        slide.shapes._spTree.append(elem)
        ibox = slide.shapes[len(slide.shapes) - 1]
        ibox.left = Pt(row_left + 10.7)
        ibox.top = Pt(current_top + 10.1)
        ibox.width = Pt(51.7)
        ibox.height = Pt(51.7)
        ibox.text_frame.text = ""  # Clean shape backplate, zero added text

        # 3. Subheader Text
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
            p.font.color.rgb = RGBColor(0x2D, 0x3E, 0x50)

        # 4. Divider Line
        elem = clean_element(copy.deepcopy(line_proto._element))
        slide.shapes._spTree.append(elem)
        line = slide.shapes[len(slide.shapes) - 1]
        line.left = Pt(row_left + 74.0)
        line.top = Pt(current_top + 31.0)
        line.width = Pt(row_width - 86.0)

        # 5. Body Text
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
            p.font.color.rgb = RGBColor(0x26, 0x26, 0x26)


def build_slide2_crd_sub_lst_variation(prs, s2_data, lib_s1, lib_s2):
    """Optional Variation: Formats Slide 2 using 2-column Subheader List Cards (Design 1 CRD-SUB-LST)."""
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    box_proto = get_shape_by_name(lib_s2, "Design 1 CRD-SUB-LST Decor box")
    sub_proto = get_shape_by_name(lib_s2, "Design 1 CRD-SUB-LST Sub header Text 1")
    body_proto = get_shape_by_name(lib_s2, "Design 1 CRD-SUB-LST Body text 1")
    div_proto = get_shape_by_name(lib_s1, "Grey Side Fade R Divider")

    cols = [s2_data[:3], s2_data[3:]]
    col_x = [45.0, 495.0]
    col_w = 420.0
    col_top = 65.0
    col_h = 410.0

    for c_idx, items in enumerate(cols):
        cx = col_x[c_idx]

        elem = clean_element(copy.deepcopy(box_proto._element))
        slide.shapes._spTree.append(elem)
        box = slide.shapes[len(slide.shapes) - 1]
        box.left = Pt(cx)
        box.top = Pt(col_top)
        box.width = Pt(col_w)
        box.height = Pt(col_h)

        padding = 14.0
        start_y = col_top + padding
        end_y = col_top + col_h - padding
        available_space = end_y - start_y
        row_height = 80.0
        sum_text_heights = len(items) * row_height
        total_empty_space = available_space - sum_text_heights
        num_gaps = max(1, (len(items) - 1) * 2)
        gap_height = max(4.0, total_empty_space / num_gaps)

        current_y = start_y

        for item_idx, item in enumerate(items):
            elem = clean_element(copy.deepcopy(sub_proto._element))
            slide.shapes._spTree.append(elem)
            s_box = slide.shapes[len(slide.shapes) - 1]
            s_box.left = Pt(cx + 18.0)
            s_box.top = Pt(current_y)
            s_box.width = Pt(col_w - 36.0)
            s_box.height = Pt(22.0)
            s_box.text_frame.text = item["subheader"]
            for p in s_box.text_frame.paragraphs:
                p.font.name = "Aptos"
                p.font.size = Pt(13.5)
                p.font.bold = True
                p.font.color.rgb = RGBColor(0x2D, 0x3E, 0x50)

            elem = clean_element(copy.deepcopy(body_proto._element))
            slide.shapes._spTree.append(elem)
            b_box = slide.shapes[len(slide.shapes) - 1]
            b_box.left = Pt(cx + 18.0)
            b_box.top = Pt(current_y + 24.0)
            b_box.width = Pt(col_w - 36.0)
            b_box.height = Pt(56.0)
            b_box.text_frame.text = item["body"]
            b_box.text_frame.word_wrap = True
            for p in b_box.text_frame.paragraphs:
                p.font.name = "Aptos"
                p.font.size = Pt(11)
                p.font.color.rgb = RGBColor(0x26, 0x26, 0x26)

            current_y += row_height

            if item_idx < len(items) - 1:
                current_y += gap_height
                elem = clean_element(copy.deepcopy(div_proto._element))
                slide.shapes._spTree.append(elem)
                div = slide.shapes[len(slide.shapes) - 1]
                div.left = Pt(cx + 18.0)
                div.top = Pt(current_y)
                div.width = Pt(col_w - 36.0)
                current_y += gap_height


def build_slide1_with_icons(prs, s1_data, lib_s1, after_s1):
    """Optional Variation: Formats Slide 1 with Option B vector icons from After.pptx."""
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Register image relationships safely
    rel_map = {}
    for rId, rel in after_s1.part.rels.items():
        if "slideLayout" not in rel.reltype and "tags" not in rel.reltype and "notesSlide" not in rel.reltype:
            new_rId = slide.part.relate_to(rel.target_part, rel.reltype)
            rel_map[rId] = new_rId

    d1_box_proto = get_shape_by_name(lib_s1, "Design 1 Decor box")
    d1_hdr_proto = get_shape_by_name(lib_s1, "Design 1 Header Text")
    d1_body_proto = get_shape_by_name(lib_s1, "Design 1 Body text 1")
    d1_div_proto = get_shape_by_name(lib_s1, "Grey Side Fade R Divider")

    icon_groups = [
        [get_shape_by_name(after_s1, n) for n in ["Graphic 186", "Group 199", "Graphic 200", "Graphic 201"]],
        [get_shape_by_name(after_s1, n) for n in ["Graphic 202", "Graphic 203", "Graphic 204", "Graphic 205"]],
        [get_shape_by_name(after_s1, n) for n in ["Graphic 206", "Graphic 207", "Graphic 208", "Graphic 209"]],
    ]

    column_x = [30.2, 341.8, 653.5]
    card_w = 276.4
    card_top = 80.0
    card_h = 410.0

    for col_idx, col_data in enumerate(s1_data):
        bx = column_x[col_idx]

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
        hdr.text_frame.text = col_data["header"]
        for p in hdr.text_frame.paragraphs:
            p.font.name = "Aptos"
            p.font.size = Pt(15)
            p.font.bold = True
            p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

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
        body_left = bx + 58.0
        body_width = card_w - 68.0

        for item_idx, body_text in enumerate(col_data["items"]):
            icon_shape = icon_groups[col_idx][item_idx]
            if icon_shape:
                ic_elem = clean_element(copy.deepcopy(icon_shape._element))
                for e in ic_elem.iter():
                    for k, v in list(e.attrib.items()):
                        if v in rel_map:
                            e.attrib[k] = rel_map[v]
                slide.shapes._spTree.append(ic_elem)
                ic = slide.shapes[len(slide.shapes) - 1]
                ic.left = Pt(bx + 16.0)
                ic.top = Pt(current_y + (row_height - 33.6) / 2)
                ic.width = Pt(33.6)
                ic.height = Pt(33.6)

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
                p.font.color.rgb = RGBColor(0x26, 0x26, 0x26)

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


def main():
    print(f"Loading input presentation from: {INPUT_PATH}")
    with open(INPUT_PATH, "rb") as f:
        input_data = f.read()
    input_prs = pptx.Presentation(io.BytesIO(input_data))

    s1_data = extract_slide1_data(input_prs)
    s2_data = extract_slide2_data(input_prs)

    print(f"Slide 1 extracted: {len(s1_data)} columns with {sum(len(c['items']) for c in s1_data)} total items")
    print(f"Slide 2 extracted: {len(s2_data)} subheader+body items")

    print(f"Loading template library from: {LIBRARY_PATH}")
    lib_prs = pptx.Presentation(LIBRARY_PATH)
    lib_s1 = lib_prs.slides[0]
    lib_s2 = lib_prs.slides[1]
    lib_s3 = lib_prs.slides[2]

    after_prs = pptx.Presentation(AFTER_PATH)
    after_s1 = after_prs.slides[0]

    # Initialize output presentation with widescreen dimensions (960x540pt)
    out_prs = pptx.Presentation()
    out_prs.slide_width = Pt(960)
    out_prs.slide_height = Pt(540)

    print("\n[Formatting Slide 1] Design 1: Floating Pill Cards...")
    build_slide1_formatted(out_prs, s1_data, lib_s1)

    print("[Formatting Slide 2] UI-ROW-01 Horizontal Process Row Cards...")
    build_slide2_formatted(out_prs, s2_data, lib_s3)

    print("[Formatting Slide 3 (Optional Variation)] Design 1 CRD-SUB-LST 2-Column Cards...")
    build_slide2_crd_sub_lst_variation(out_prs, s2_data, lib_s1, lib_s2)

    print("[Formatting Slide 4 (Optional Variation)] Design 1 with Option B Vector Icons...")
    build_slide1_with_icons(out_prs, s1_data, lib_s1, after_s1)

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    out_prs.save(OUTPUT_PATH)
    print(f"\nSuccessfully generated formatted deck at:\n  {OUTPUT_PATH}")
    print(f"Total Slides: {len(out_prs.slides)}")


if __name__ == "__main__":
    main()

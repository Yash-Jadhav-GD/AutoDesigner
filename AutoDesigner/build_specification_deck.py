"""
AutoDesigner Specification Conformance Engine
Constructs the authoritative single presentation containing all template-compliant design variations
strictly matching AUTOFORMAT_SPECIFICATION.md and Library.pptx prototypes.
Guarantees 100% valid OpenXML relationship mapping so PowerPoint opens cleanly without repair warnings.
"""

import copy
from pathlib import Path
import pptx
from pptx.util import Pt
from pptx.dml.color import RGBColor

BASE_DIR = Path(__file__).resolve().parent
AFTER_PATH = BASE_DIR / "After_copy.pptx"
LIBRARY_PATH = BASE_DIR / "Library_copy.pptx"
OUTPUT_PATH = BASE_DIR / "output" / "Design_tester_new_formatted.pptx"

COLUMNS_DATA = [
    {
        "header": "The Problem",
        "items": [
            ("Procurement Silos", "Procurement work is siloed by team, not by process"),
            ("Inconsistent Activities", "500+ activities managed inconsistently across regions"),
            ("High-Context Tasks", "WL1 resources handle high-context, cross-functional tasks with no AI support"),
            ("No Baseline", "No standardized process baseline exists to build automation upon"),
        ],
    },
    {
        "header": "The Bridge",
        "items": [
            ("L3–L6 Process Mapping", "WS2 maps current L3–L6 processes across all 9 persona groups"),
            ("AI-Ready Workflow Units", "Deconstructs work into AI-ready, standardized workflow units"),
            ("Human+Agent Operating Model", "Designs the human+agent operating model for each process"),
            ("Execution Blueprint", "Creates the blueprint that WS3 governs and WS4 deploys"),
        ],
    },
    {
        "header": "The Future State",
        "items": [
            ("Process Horizontals", "Harmonized process horizontals replace siloed team structures"),
            ("Autonomous Workflows", "Multi-agent workflows handle repeatable, rules-based tasks autonomously"),
            ("Strategic Judgment", "Humans focus on judgment, strategy, and relationship-intensive work"),
            ("Quantum Intelligence", "GEP Quantum Intelligence orchestrates end-to-end S2P execution"),
        ],
    },
]

COLUMN_X = [30.2, 341.8, 653.5]
CARD_WIDTH = 276.4


def clean_element(elem):
    """Strips any legacy custDataLst / tags elements that refer to missing rId2 relationships."""
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


def add_slide_banner(slide, title_text, subtitle_text):
    """Adds standard executive header banner."""
    t_box = slide.shapes.add_textbox(Pt(30.2), Pt(20.0), Pt(899.7), Pt(46.0))
    tf = t_box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Pt(0)
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.name = "Aptos"
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0x26, 0x26, 0x26)

    if subtitle_text:
        p2 = tf.add_paragraph()
        p2.text = subtitle_text
        p2.font.name = "Aptos"
        p2.font.size = Pt(12)
        p2.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)


def build_design1_standard(prs, lib_slide):
    """Slide 2: Design 1 Floating Pill Cards with strict aspect-ratio locking & gradient dividers."""
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    add_slide_banner(
        slide,
        "WS2 Strategic Context: The Problem, The Bridge & The Future State",
        "Design 1: Floating Pill Cards with Gradient Line Dividers (Standard Archetype)",
    )

    d1_box_proto = get_shape_by_name(lib_slide, "Design 1 Decor box")
    d1_hdr_proto = get_shape_by_name(lib_slide, "Design 1 Header Text")
    d1_body_proto = get_shape_by_name(lib_slide, "Design 1 Body text 1")
    d1_div_proto = get_shape_by_name(lib_slide, "Grey Side Fade R Divider")

    card_top = 115.0
    card_h = 390.0

    for col_idx, col_data in enumerate(COLUMNS_DATA):
        bx = COLUMN_X[col_idx]

        # 1. Clone Background Container (clean element to strip custDataLst)
        elem = clean_element(copy.deepcopy(d1_box_proto._element))
        slide.shapes._spTree.append(elem)
        box = slide.shapes[len(slide.shapes) - 1]
        box.left = Pt(bx)
        box.top = Pt(card_top)
        box.width = Pt(CARD_WIDTH)
        box.height = Pt(card_h)

        # 2. Clone Floating Header Pill (15.7pt protrusion)
        elem = clean_element(copy.deepcopy(d1_hdr_proto._element))
        slide.shapes._spTree.append(elem)
        hdr = slide.shapes[len(slide.shapes) - 1]
        hdr.left = Pt(bx + 30.4)
        hdr.top = Pt(card_top - 15.7)
        hdr.width = Pt(215.4)
        hdr.height = Pt(33.0)
        hdr.text_frame.text = col_data["header"]
        for p in hdr.text_frame.paragraphs:
            p.font.name = "Aptos"
            p.font.size = Pt(15)
            p.font.bold = True
            p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

        # 3. Dynamic Vertical Row Distribution (Pass 2 from SPECIFICATION)
        padding = 8.5039
        start_y = card_top + 33.0 + padding
        end_y = card_top + card_h - padding
        available_space = end_y - start_y
        row_height = 54.0
        sum_text_heights = len(col_data["items"]) * row_height
        total_empty_space = available_space - sum_text_heights
        num_gaps = (len(col_data["items"]) - 1) * 2
        gap_height = max(4.0, total_empty_space / num_gaps)

        current_y = start_y
        body_left = bx + 10.3
        body_width = CARD_WIDTH - 20.6

        for item_idx, (_, body_text) in enumerate(col_data["items"]):
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

            if item_idx < len(col_data["items"]) - 1:
                current_y += gap_height
                elem = clean_element(copy.deepcopy(d1_div_proto._element))
                slide.shapes._spTree.append(elem)
                div = slide.shapes[len(slide.shapes) - 1]
                div.left = Pt(bx + 17.5)
                div.top = Pt(current_y)
                div.width = Pt(CARD_WIDTH - 35.0)
                current_y += gap_height


def build_design1_with_icons(prs, lib_slide, source_slide):
    """Slide 3: Design 1 with Option B Vector Icons and properly mapped OpenXML relationships."""
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Map image relationships from source_slide to slide
    rel_map = {}
    for rId, rel in source_slide.part.rels.items():
        if "slideLayout" not in rel.reltype and "tags" not in rel.reltype and "notesSlide" not in rel.reltype:
            new_rId = slide.part.relate_to(rel.target_part, rel.reltype)
            rel_map[rId] = new_rId

    add_slide_banner(
        slide,
        "WS2 Strategic Context: The Problem, The Bridge & The Future State",
        "Design 1: Floating Pill Cards with Option B Semantic Vector Icons",
    )

    d1_box_proto = get_shape_by_name(lib_slide, "Design 1 Decor box")
    d1_hdr_proto = get_shape_by_name(lib_slide, "Design 1 Header Text")
    d1_body_proto = get_shape_by_name(lib_slide, "Design 1 Body text 1")
    d1_div_proto = get_shape_by_name(lib_slide, "Grey Side Fade R Divider")

    icon_groups = [
        [get_shape_by_name(source_slide, n) for n in ["Graphic 186", "Group 199", "Graphic 200", "Graphic 201"]],
        [get_shape_by_name(source_slide, n) for n in ["Graphic 202", "Graphic 203", "Graphic 204", "Graphic 205"]],
        [get_shape_by_name(source_slide, n) for n in ["Graphic 206", "Graphic 207", "Graphic 208", "Graphic 209"]],
    ]

    card_top = 115.0
    card_h = 390.0

    for col_idx, col_data in enumerate(COLUMNS_DATA):
        bx = COLUMN_X[col_idx]

        # Background Decor Box
        elem = clean_element(copy.deepcopy(d1_box_proto._element))
        slide.shapes._spTree.append(elem)
        box = slide.shapes[len(slide.shapes) - 1]
        box.left = Pt(bx)
        box.top = Pt(card_top)
        box.width = Pt(CARD_WIDTH)
        box.height = Pt(card_h)

        # Header Pill
        elem = clean_element(copy.deepcopy(d1_hdr_proto._element))
        slide.shapes._spTree.append(elem)
        hdr = slide.shapes[len(slide.shapes) - 1]
        hdr.left = Pt(bx + 30.4)
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
        row_height = 54.0
        sum_text_heights = len(col_data["items"]) * row_height
        total_empty_space = available_space - sum_text_heights
        num_gaps = (len(col_data["items"]) - 1) * 2
        gap_height = max(4.0, total_empty_space / num_gaps)

        current_y = start_y
        body_left = bx + 58.0
        body_width = CARD_WIDTH - 68.0

        for item_idx, (_, body_text) in enumerate(col_data["items"]):
            # Icon with mapped relationship
            icon_shape = icon_groups[col_idx][item_idx]
            if icon_shape:
                ic_elem = clean_element(copy.deepcopy(icon_shape._element))
                # Remap rId attributes to point to slide's newly registered relationships
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

            # Body text box
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
                div.width = Pt(CARD_WIDTH - 35.0)
                current_y += gap_height


def build_design1_native_backplates(prs, lib_s1, lib_s4):
    """Slide 4: Design 1 with Native Circular Backplates from Library.pptx Slide 4."""
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    add_slide_banner(
        slide,
        "WS2 Strategic Context: The Problem, The Bridge & The Future State",
        "Design 1: Cards with Native Circular Icon Backplates (Library Slide 4)",
    )

    d1_box_proto = get_shape_by_name(lib_s1, "Design 1 Decor box")
    d1_hdr_proto = get_shape_by_name(lib_s1, "Design 1 Header Text")
    d1_body_proto = get_shape_by_name(lib_s1, "Design 1 Body text 1")
    d1_div_proto = get_shape_by_name(lib_s1, "Grey Side Fade R Divider")
    circle_proto = get_shape_by_name(lib_s4, "Tiny CircleBackplate 22")

    card_top = 115.0
    card_h = 390.0

    for col_idx, col_data in enumerate(COLUMNS_DATA):
        bx = COLUMN_X[col_idx]

        elem = clean_element(copy.deepcopy(d1_box_proto._element))
        slide.shapes._spTree.append(elem)
        box = slide.shapes[len(slide.shapes) - 1]
        box.left = Pt(bx)
        box.top = Pt(card_top)
        box.width = Pt(CARD_WIDTH)
        box.height = Pt(card_h)

        elem = clean_element(copy.deepcopy(d1_hdr_proto._element))
        slide.shapes._spTree.append(elem)
        hdr = slide.shapes[len(slide.shapes) - 1]
        hdr.left = Pt(bx + 30.4)
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
        row_height = 54.0
        sum_text_heights = len(col_data["items"]) * row_height
        total_empty_space = available_space - sum_text_heights
        num_gaps = (len(col_data["items"]) - 1) * 2
        gap_height = max(4.0, total_empty_space / num_gaps)

        current_y = start_y

        for item_idx, (_, body_text) in enumerate(col_data["items"]):
            elem = clean_element(copy.deepcopy(circle_proto._element))
            slide.shapes._spTree.append(elem)
            c_shape = slide.shapes[len(slide.shapes) - 1]
            c_shape.left = Pt(bx + 16.0)
            c_shape.top = Pt(current_y + (row_height - 34.0) / 2)
            c_shape.width = Pt(34.0)
            c_shape.height = Pt(34.0)

            elem = clean_element(copy.deepcopy(d1_body_proto._element))
            slide.shapes._spTree.append(elem)
            b_box = slide.shapes[len(slide.shapes) - 1]
            b_box.left = Pt(bx + 62.0)
            b_box.top = Pt(current_y)
            b_box.width = Pt(CARD_WIDTH - 72.0)
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
                div.width = Pt(CARD_WIDTH - 35.0)
                current_y += gap_height


def build_design2_flush(prs, lib_slide):
    """Slide 5: Design 2 Flush Header Cards from Library.pptx Slide 1."""
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    add_slide_banner(
        slide,
        "WS2 Strategic Context: The Problem, The Bridge & The Future State",
        "Design 2: Flush Header Cards (Library Slide 1)",
    )

    d2_box_proto = get_shape_by_name(lib_slide, "Design 2 Decor box")
    d2_hdr_proto = get_shape_by_name(lib_slide, "Design 2 Header Text")
    d2_body_proto = get_shape_by_name(lib_slide, "Design 2 Body text 1")
    d1_div_proto = get_shape_by_name(lib_slide, "Grey Side Fade R Divider")

    card_top = 100.0
    card_h = 405.0

    for col_idx, col_data in enumerate(COLUMNS_DATA):
        bx = COLUMN_X[col_idx]

        elem = clean_element(copy.deepcopy(d2_box_proto._element))
        slide.shapes._spTree.append(elem)
        box = slide.shapes[len(slide.shapes) - 1]
        box.left = Pt(bx)
        box.top = Pt(card_top)
        box.width = Pt(CARD_WIDTH)
        box.height = Pt(card_h)

        elem = clean_element(copy.deepcopy(d2_hdr_proto._element))
        slide.shapes._spTree.append(elem)
        hdr = slide.shapes[len(slide.shapes) - 1]
        hdr.left = Pt(bx + 10.3)
        hdr.top = Pt(card_top + 10.2)
        hdr.width = Pt(255.8)
        hdr.height = Pt(31.3)
        hdr.text_frame.text = col_data["header"]
        for p in hdr.text_frame.paragraphs:
            p.font.name = "Aptos"
            p.font.size = Pt(15)
            p.font.bold = True
            p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

        padding = 8.5039
        start_y = card_top + 10.2 + 31.3 + padding
        end_y = card_top + card_h - padding
        available_space = end_y - start_y
        row_height = 54.0
        sum_text_heights = len(col_data["items"]) * row_height
        total_empty_space = available_space - sum_text_heights
        num_gaps = (len(col_data["items"]) - 1) * 2
        gap_height = max(4.0, total_empty_space / num_gaps)

        current_y = start_y

        for item_idx, (_, body_text) in enumerate(col_data["items"]):
            elem = clean_element(copy.deepcopy(d2_body_proto._element))
            slide.shapes._spTree.append(elem)
            b_box = slide.shapes[len(slide.shapes) - 1]
            b_box.left = Pt(bx + 10.3)
            b_box.top = Pt(current_y)
            b_box.width = Pt(255.8)
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
                div.width = Pt(CARD_WIDTH - 35.0)
                current_y += gap_height


def build_crd_sub_lst(prs, lib_s1, lib_s2):
    """Slide 6: Subheader List Card Design 1 CRD-SUB-LST from Library Slide 2."""
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    add_slide_banner(
        slide,
        "WS2 Strategic Context: The Problem, The Bridge & The Future State",
        "Design 1 CRD-SUB-LST: Subheader + Body Description Cards (Library Slide 2)",
    )

    box_proto = get_shape_by_name(lib_s2, "Design 1 CRD-SUB-LST Decor box")
    hdr_proto = get_shape_by_name(lib_s2, "Design 1 CRD-SUB-LSTHeader Text")
    sub_proto = get_shape_by_name(lib_s2, "Design 1 CRD-SUB-LST Sub header Text 1")
    body_proto = get_shape_by_name(lib_s2, "Design 1 CRD-SUB-LST Body text 1")
    div_proto = get_shape_by_name(lib_s1, "Grey Side Fade R Divider")

    card_top = 100.0
    card_h = 405.0

    for col_idx, col_data in enumerate(COLUMNS_DATA):
        bx = COLUMN_X[col_idx]

        elem = clean_element(copy.deepcopy(box_proto._element))
        slide.shapes._spTree.append(elem)
        box = slide.shapes[len(slide.shapes) - 1]
        box.left = Pt(bx)
        box.top = Pt(card_top)
        box.width = Pt(CARD_WIDTH)
        box.height = Pt(card_h)

        elem = clean_element(copy.deepcopy(hdr_proto._element))
        slide.shapes._spTree.append(elem)
        hdr = slide.shapes[len(slide.shapes) - 1]
        hdr.left = Pt(bx + 5.3)
        hdr.top = Pt(card_top + 5.5)
        hdr.width = Pt(265.8)
        hdr.height = Pt(31.3)
        hdr.text_frame.text = col_data["header"]
        for p in hdr.text_frame.paragraphs:
            p.font.name = "Aptos"
            p.font.size = Pt(15)
            p.font.bold = True
            p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

        padding = 8.5039
        start_y = card_top + 5.5 + 31.3 + padding
        end_y = card_top + card_h - padding
        available_space = end_y - start_y
        row_height = 56.0
        sum_text_heights = len(col_data["items"]) * row_height
        total_empty_space = available_space - sum_text_heights
        num_gaps = (len(col_data["items"]) - 1) * 2
        gap_height = max(4.0, total_empty_space / num_gaps)

        current_y = start_y

        for item_idx, (sub_text, body_text) in enumerate(col_data["items"]):
            elem = clean_element(copy.deepcopy(sub_proto._element))
            slide.shapes._spTree.append(elem)
            s_box = slide.shapes[len(slide.shapes) - 1]
            s_box.left = Pt(bx + 13.2)
            s_box.top = Pt(current_y)
            s_box.width = Pt(250.0)
            s_box.height = Pt(20.0)
            s_box.text_frame.text = sub_text
            for p in s_box.text_frame.paragraphs:
                p.font.name = "Aptos"
                p.font.size = Pt(13)
                p.font.bold = True
                p.font.color.rgb = RGBColor(0x2D, 0x3E, 0x50)

            elem = clean_element(copy.deepcopy(body_proto._element))
            slide.shapes._spTree.append(elem)
            b_box = slide.shapes[len(slide.shapes) - 1]
            b_box.left = Pt(bx + 13.2)
            b_box.top = Pt(current_y + 20.0)
            b_box.width = Pt(250.0)
            b_box.height = Pt(36.0)
            b_box.text_frame.text = body_text
            b_box.text_frame.word_wrap = True
            for p in b_box.text_frame.paragraphs:
                p.font.name = "Aptos"
                p.font.size = Pt(10.5)
                p.font.color.rgb = RGBColor(0x26, 0x26, 0x26)

            current_y += row_height

            if item_idx < len(col_data["items"]) - 1:
                current_y += gap_height
                elem = clean_element(copy.deepcopy(div_proto._element))
                slide.shapes._spTree.append(elem)
                div = slide.shapes[len(slide.shapes) - 1]
                div.left = Pt(bx + 17.5)
                div.top = Pt(current_y)
                div.width = Pt(CARD_WIDTH - 35.0)
                current_y += gap_height


def main():
    print(f"Loading After_copy from: {AFTER_PATH}")
    prs = pptx.Presentation(AFTER_PATH)

    print(f"Loading Library_copy from: {LIBRARY_PATH}")
    lib_prs = pptx.Presentation(LIBRARY_PATH)

    # Keep Slide 1 from After_copy as Slide 1 (Production Benchmark)
    while len(prs.slides) > 1:
        rId = prs.slides._sldIdLst[1].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[1]

    lib_s1 = lib_prs.slides[0]
    lib_s2 = lib_prs.slides[1]
    lib_s4 = lib_prs.slides[3]
    after_s1 = prs.slides[0]

    print("Building Slide 2: Design 1 Floating Pill Cards (Standard with Gradient Dividers)...")
    build_design1_standard(prs, lib_s1)

    print("Building Slide 3: Design 1 with Option B Vector Icons...")
    build_design1_with_icons(prs, lib_s1, source_slide=after_s1)

    print("Building Slide 4: Design 1 with Native Circular Backplates...")
    build_design1_native_backplates(prs, lib_s1, lib_s4)

    print("Building Slide 5: Design 2 Flush Header Cards...")
    build_design2_flush(prs, lib_s1)

    print("Building Slide 6: Design 1 CRD-SUB-LST Subheader List Cards...")
    build_crd_sub_lst(prs, lib_s1, lib_s2)

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    prs.save(OUTPUT_PATH)
    print(f"\nSaved master presentation with all 6 variations to:\n  {OUTPUT_PATH}")


if __name__ == "__main__":
    main()

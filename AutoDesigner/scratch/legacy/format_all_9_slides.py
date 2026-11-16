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
from engine.builders.common import set_cell_border, add_connector_line
from engine.color_tokens import (
    ColorTokens,
    hex_to_rgb,
    apply_linear_gradient,
    apply_solid_fill,
)

INPUT_PATH = BASE_DIR / "Design tester new.pptx"
INPUT_INSPECT_PATH = BASE_DIR / "Design_tester_new_inspect.pptx"
AFTER_PATH = BASE_DIR / "After_copy.pptx"
LIBRARY_PATH = BASE_DIR / "Library_copy.pptx"
LIBRARY_INSPECT_PATH = BASE_DIR / "Library_inspect.pptx"
OUTPUT_PATH = BASE_DIR / "output" / "Design_tester_new_formatted.pptx"
LOCAL_COPY_PATH = BASE_DIR / "Design tester new_formatted.pptx"

# Import builders 1-6 from format_all_6_slides
import format_all_6_slides as f6



def extract_slide7_data(prs):
    return {
        "title": "Delivery Team Structure: WL2 Leadership Archetypes",
        "top_banner": "PaaS 2.0 Program Management Team",
        "leadership": [
            {
                "title": "Tier 2/3 Value-Creation Category Leads",
                "sub": None,
                "ftes": "31 FTEs",
                "type": "spend_value",
            },
            {
                "title": "Tier 1 Spend Execution Leads",
                "sub": None,
                "ftes": "10 FTEs",
                "type": "span_breaker",
            },
            {
                "title": "BU/BG Leads",
                "sub": None,
                "ftes": "9 FTEs",
                "type": "span_breaker",
            },
            {
                "title": "Functional Leads",
                "sub": "(Sustainability, S&I and POC)",
                "ftes": "2 FTEs",
                "type": "span_breaker",
            },
        ],
        "pool1": {
            "title": "Portfolio WL1 Delivery Pool",
            "badge": "240 - 250 WL1 roles",
            "sub": "Shared pool deployed across both Tier 1 execution and Tier 2/3 value creation – the same portfolio departments carry spend in both tiers",
            "cards": [
                "Portfolio – RPM",
                "Portfolio – IT/BS",
                "Portfolio – Marketing",
                "Portfolio – CO",
                "Capex & Utilities",
                "Real Estate",
            ],
        },
        "pool2": {
            "title": "BU/BG WL1 Teams",
            "badge": "150 - 160 WL1 roles",
            "sub": "BU- and BG-dedicated teams, deployed across all eight regions and reporting into the BU/BG leads",
            "regions": [
                "North America",
                "Europe",
                "PTAB",
                "LATAM",
                "Greater Asia",
                "Greater China",
                "Indonesia",
                "India",
            ],
        },
        "pool3": {
            "title": "Functional WL1 Teams",
            "badge": "50 - 60 WL1 roles",
            "sub": "Held as discrete lines as functional agendas are not absorbed into portfolio or BU/BG delivery priorities.",
            "cards": [
                "POC",
                "S&I",
                "Sustainability",
            ],
        },
        "legend": {
            "l1": "Spend and Value Focused",
            "l2": "Span Breakers",
        },
    }


def extract_slide8_data(prs):
    s8 = prs.slides[7]
    title = None
    note = None
    footer = None
    tbl_shape = None
    for shp in s8.shapes:
        if shp.has_table:
            tbl_shape = shp
        elif shp.has_text_frame:
            txt = shp.text_frame.text.strip()
            if "Spend and Value Based" in txt:
                title = txt
            elif "Packaging sub-categories" in txt:
                footer = txt
            elif "Considered €200 MN" in txt:
                note = txt
        elif shp.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.GROUP:
            for gs in shp.shapes:
                if gs.has_text_frame:
                    gtxt = gs.text_frame.text.strip()
                    if "Considered €200 MN" in gtxt:
                        note = gtxt

    table_data = []
    for r in tbl_shape.table.rows:
        row_cells = []
        for c in r.cells:
            p_list = [p.text.strip() for p in c.text_frame.paragraphs if p.text.strip()]
            row_cells.append(p_list)
        table_data.append(row_cells)

    return {
        "title": title,
        "note": note,
        "footer": footer,
        "table_data": table_data,
    }


def extract_slide9_data(prs):
    s9 = prs.slides[8]
    title = None
    tbl_shape = None
    for shp in s9.shapes:
        if shp.has_table:
            tbl_shape = shp
        elif shp.has_text_frame:
            txt = shp.text_frame.text.strip()
            if "Span of Control Driven" in txt:
                title = txt

    table_data = []
    for r in tbl_shape.table.rows:
        row_cells = []
        for c in r.cells:
            p_list = [p.text.strip() for p in c.text_frame.paragraphs if p.text.strip()]
            row_cells.append(p_list)
        table_data.append(row_cells)

    return {
        "title": title,
        "table_data": table_data,
    }




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

    f6.add_library_line(slide, lib_lines["side_fade_r"], 44.0, pool_y + 66.0, 412.0)

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

    f6.add_library_line(slide, lib_lines["side_fade_r"], 500.0, pool_y + 84.0, 190.0)

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

    f6.add_library_line(slide, lib_lines["side_fade_r"], 730.0, pool_y + 84.0, 190.0)

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


def main():
    print(f"Loading input presentation: {INPUT_INSPECT_PATH}")
    with open(INPUT_INSPECT_PATH, "rb") as f:
        input_data = f.read()
    input_prs = pptx.Presentation(io.BytesIO(input_data))

    s1_data = f6.extract_slide1_data(input_prs)
    s2_data = f6.extract_slide2_data(input_prs)
    s3_data = f6.extract_slide3_data(input_prs)
    s4_data = f6.extract_slide4_data(input_prs)
    s5_data = f6.extract_slide5_data(input_prs)
    s6_data = f6.extract_slide6_data(input_prs)
    s7_data = extract_slide7_data(input_prs)
    s8_data = extract_slide8_data(input_prs)
    s9_data = extract_slide9_data(input_prs)

    print(f"Extracted all 9 slides successfully!")

    # Safely load Library Presentation
    lib_path_to_use = LIBRARY_PATH
    if not lib_path_to_use.exists() or os.access(lib_path_to_use, os.R_OK) is False:
        lib_path_to_use = LIBRARY_INSPECT_PATH
    try:
        with open(lib_path_to_use, "rb") as f:
            lib_bytes = f.read()
        lib_prs = pptx.Presentation(io.BytesIO(lib_bytes))
    except Exception:
        lib_prs = pptx.Presentation(LIBRARY_INSPECT_PATH)

    lib_s1 = lib_prs.slides[0]
    lib_s3 = lib_prs.slides[2]
    lib_s6 = lib_prs.slides[5]

    lib_lines = {
        "side_fade_r": f6.get_shape_by_name(lib_s6, "Grey Side Fade R"),
        "centre_fade_h": f6.get_shape_by_name(lib_s6, "Grey Centre Fade H"),
        "solid_h": f6.get_shape_by_name(lib_s6, "Grey Solid H"),
        "solid_v": f6.get_shape_by_name(lib_s6, "Grey Solid V"),
        "centre_fade_v": f6.get_shape_by_name(lib_s6, "Grey Centre Fade V"),
    }

    out_prs = pptx.Presentation()
    out_prs.slide_width = Pt(960)
    out_prs.slide_height = Pt(540)

    print("\n[Formatting Slide 1] 3-Column Narrative...")
    f6.build_slide1(out_prs, s1_data, lib_s1, lib_lines)

    print("[Formatting Slide 2] Split Layout with SVG Icons...")
    f6.build_slide2(out_prs, s2_data, lib_s1, lib_s3, lib_lines)

    print("[Formatting Slide 3] Full-Width Process Rows...")
    f6.build_slide3(out_prs, s3_data, lib_s3, lib_lines)

    print("[Formatting Slide 4] Phase 1 Assessment with SVG Icons & Key Insights...")
    f6.build_slide4(out_prs, s4_data, lib_s1, lib_s3, lib_lines)

    print("[Formatting Slide 5] Phase 3 Five Multi-Agent Workflows...")
    f6.build_slide5(out_prs, s5_data, lib_s3, lib_lines)

    print("[Formatting Slide 6] WS2 Team Structure 3x3 Grid...")
    f6.build_slide6(out_prs, s6_data, lib_s3, lib_lines)

    print("[Formatting Slide 7] Delivery Team Structure: WL2 Leadership Archetypes...")
    build_slide7(out_prs, s7_data, lib_lines)

    print("[Formatting Slide 8] Spend & Value Based WL2 Allocation (Table)...")
    build_slide8(out_prs, s8_data)

    print("[Formatting Slide 9] Span of Control Driven WL2 Sizing (Table)...")
    build_slide9(out_prs, s9_data)

    print(f"Total slides built: {len(out_prs.slides)}")
    assert len(out_prs.slides) == 9, "Must have exactly 9 slides!"

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    buf = io.BytesIO()
    out_prs.save(buf)
    raw_bytes = buf.getvalue()

    print("[Enriching PPTX Package] Embedding SVG vector icons and updating relationships...")
    file_bytes = f6.inject_svg_assets(raw_bytes, AFTER_PATH)

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

    print(f"\nSuccessfully generated formatted presentations with exactly 9 slides at:")
    for sp in saved_paths:
        print(f"  - {sp}")


if __name__ == "__main__":
    main()

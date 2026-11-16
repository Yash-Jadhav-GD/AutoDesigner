"""
Extractors that turn the known "Design tester new" draft deck into JSON-serialisable
slide data. Used once to seed spec files; new decks should be authored as specs directly.
"""

import pptx
import pptx.enum.shapes


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

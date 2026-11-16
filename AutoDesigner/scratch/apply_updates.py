import re
from pathlib import Path

target = Path(r"c:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\scratch\builders_8_to_20.py")
content = target.read_text(encoding="utf-8")

# 1. Slide 17 bullet replacements
content = content.replace(
    '("L3", "Service Director • Delivery Manager • Category & Commercial Managers", ColorTokens.BRAND_NAVY)',
    '("L3", "Service Director | Delivery Manager | Category & Commercial Managers", ColorTokens.BRAND_NAVY)'
)
content = content.replace(
    '("L2", "Sourcing & Contracting Specialists • Governance • Materials • Local Content", ColorTokens.BRAND_INDIGO)',
    '("L2", "Sourcing & Contracting Specialists | Governance | Materials | Local Content", ColorTokens.BRAND_INDIGO)'
)
content = content.replace(
    '("L1", "Analysts • Buyers • Contract Administrators • Reporting", ColorTokens.BRAND_TANGERINE)',
    '("L1", "Analysts | Buyers | Contract Administrators | Reporting", ColorTokens.BRAND_TANGERINE)'
)

# 2. Slide 20 bullet replacement
content = content.replace(
    'lp1.text = "Performance Trends  •  Critical Issues  •  Long-Term Strategies  •  Project Updates  •  Critical Issues, Risks and Recommendations  •  Savings / Compliance Reports, SLA’s, Survey Results"',
    'lp1.text = "Performance Trends  |  Critical Issues  |  Long-Term Strategies  |  Project Updates  |  Critical Issues, Risks and Recommendations  |  Savings / Compliance Reports, SLA’s, Survey Results"'
)

# 3. Slide 14/15 title and subtitle spacing
content = content.replace(
    'title_box = slide.shapes.add_textbox(Pt(30.0), Pt(6.0), Pt(900.0), Pt(24.0))\n    tf = title_box.text_frame\n    tf.word_wrap = True\n    p = tf.paragraphs[0]\n    p.text = "Every interface has an owner, and issues escalate before they cause delay"\n    p.font.name = "Aptos"\n    p.font.size = Pt(14.0)\n    p.font.bold = True\n    p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)\n\n    sub_box = slide.shapes.add_textbox(Pt(30.0), Pt(28.0), Pt(900.0), Pt(18.0))',
    'title_box = slide.shapes.add_textbox(Pt(30.0), Pt(8.0), Pt(900.0), Pt(20.0))\n    tf = title_box.text_frame\n    tf.word_wrap = True\n    p = tf.paragraphs[0]\n    p.text = "Every interface has an owner, and issues escalate before they cause delay"\n    p.font.name = "Aptos"\n    p.font.size = Pt(14.0)\n    p.font.bold = True\n    p.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_ONYX)\n\n    sub_box = slide.shapes.add_textbox(Pt(30.0), Pt(30.0), Pt(900.0), Pt(16.0))'
)

# 4. Slide 14/15 Spokes coordinates
old_spokes = '''    # 10 Spokes around the hub
    spokes = [
        ("bp IJV", 45.0, 100.0),
        ("Operations", 195.0, 90.0),
        ("Wells", 335.0, 100.0),
        ("Projects", 345.0, 170.0),
        ("Engineering", 335.0, 240.0),
        ("HSE", 270.0, 310.0),
        ("Finance", 130.0, 310.0),
        ("Logistics", 45.0, 240.0),
        ("Security", 35.0, 170.0),
        ("NOC / NGC", 110.0, 95.0)
    ]'''

new_spokes = '''    # 10 Spokes around the hub (strictly disjoint coordinates)
    spokes = [
        ("bp IJV", 42.0, 96.0),
        ("Operations", 150.0, 96.0),
        ("Wells", 260.0, 96.0),
        ("Projects", 368.0, 96.0),
        ("Security", 42.0, 175.0),
        ("Engineering", 368.0, 175.0),
        ("Logistics", 42.0, 254.0),
        ("NOC / NGC", 368.0, 254.0),
        ("Finance", 150.0, 295.0),
        ("HSE", 260.0, 295.0),
    ]'''

content = content.replace(old_spokes, new_spokes)

# 5. Slide 14/15 Escalation level boxes and arrows (ey spacing)
old_esc = '''    for e_i, (lvl_num, lvl_type, lvl_leads, lvl_desc, lvl_clr) in enumerate(esc_levels):
        ey = 90.0 + e_i * 84.0
        ebox = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(498.0), Pt(ey), Pt(424.0), Pt(70.0))
        apply_solid_fill(ebox, "F8FAFC")
        ebox.line.color.rgb = hex_to_rgb("CBD5E1")
        set_shape_corner_radius(ebox, 4.0)
        apply_library_shadow(ebox)

        # Level Badge
        l_badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(506.0), Pt(ey + 8.0), Pt(74.0), Pt(54.0))
        apply_solid_fill(l_badge, lvl_clr)
        set_shape_corner_radius(l_badge, 4.0)
        ltf = l_badge.text_frame
        ltf.vertical_anchor = MSO_ANCHOR.MIDDLE
        lp1 = ltf.paragraphs[0]
        lp1.text = lvl_num
        lp1.alignment = PP_ALIGN.CENTER
        lp1.font.name = "Aptos"
        lp1.font.size = Pt(10.0)
        lp1.font.bold = True
        lp1.font.color.rgb = hex_to_rgb("FFFFFF")
        lp2 = ltf.add_paragraph()
        lp2.text = lvl_type
        lp2.alignment = PP_ALIGN.CENTER
        lp2.font.name = "Aptos"
        lp2.font.size = Pt(8.0)
        lp2.font.color.rgb = hex_to_rgb("E2E8F0")

        # Details
        dtf_box = slide.shapes.add_textbox(Pt(588.0), Pt(ey + 6.0), Pt(326.0), Pt(58.0))
        dtf = dtf_box.text_frame
        dtf.word_wrap = True
        dtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        dp1 = dtf.paragraphs[0]
        dp1.text = lvl_leads
        dp1.font.name = "Aptos"
        dp1.font.size = Pt(9.5)
        dp1.font.bold = True
        dp1.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

        dp2 = dtf.add_paragraph()
        dp2.text = lvl_desc
        dp2.font.name = "Aptos"
        dp2.font.size = Pt(8.0)
        dp2.font.color.rgb = hex_to_rgb("475569")

        # Arrow up indicator between levels
        if e_i < 2:
            arr_box = slide.shapes.add_textbox(Pt(680.0), Pt(ey + 68.0), Pt(40.0), Pt(18.0))'''

new_esc = '''    for e_i, (lvl_num, lvl_type, lvl_leads, lvl_desc, lvl_clr) in enumerate(esc_levels):
        ey = 90.0 + e_i * 84.0
        ebox = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(498.0), Pt(ey), Pt(424.0), Pt(64.0))
        apply_solid_fill(ebox, "F8FAFC")
        ebox.line.color.rgb = hex_to_rgb("CBD5E1")
        set_shape_corner_radius(ebox, 4.0)
        apply_library_shadow(ebox)

        # Level Badge
        l_badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(506.0), Pt(ey + 6.0), Pt(74.0), Pt(52.0))
        apply_solid_fill(l_badge, lvl_clr)
        set_shape_corner_radius(l_badge, 4.0)
        ltf = l_badge.text_frame
        ltf.vertical_anchor = MSO_ANCHOR.MIDDLE
        lp1 = ltf.paragraphs[0]
        lp1.text = lvl_num
        lp1.alignment = PP_ALIGN.CENTER
        lp1.font.name = "Aptos"
        lp1.font.size = Pt(10.0)
        lp1.font.bold = True
        lp1.font.color.rgb = hex_to_rgb("FFFFFF")
        lp2 = ltf.add_paragraph()
        lp2.text = lvl_type
        lp2.alignment = PP_ALIGN.CENTER
        lp2.font.name = "Aptos"
        lp2.font.size = Pt(9.0)
        lp2.font.color.rgb = hex_to_rgb("E2E8F0")

        # Details
        dtf_box = slide.shapes.add_textbox(Pt(588.0), Pt(ey + 4.0), Pt(326.0), Pt(56.0))
        dtf = dtf_box.text_frame
        dtf.word_wrap = True
        dtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        dp1 = dtf.paragraphs[0]
        dp1.text = lvl_leads
        dp1.font.name = "Aptos"
        dp1.font.size = Pt(9.5)
        dp1.font.bold = True
        dp1.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

        dp2 = dtf.add_paragraph()
        dp2.text = lvl_desc
        dp2.font.name = "Aptos"
        dp2.font.size = Pt(9.0)
        dp2.font.color.rgb = hex_to_rgb("475569")

        # Arrow up indicator between levels
        if e_i < 2:
            arr_box = slide.shapes.add_textbox(Pt(680.0), Pt(ey + 66.0), Pt(40.0), Pt(14.0))'''

content = content.replace(old_esc, new_esc)

# 6. Slide 12 Footnote & Infrastructure placement
old_s12 = '''    # Footnote
    fn_box = slide.shapes.add_textbox(Pt(30.0), Pt(390.0), Pt(900.0), Pt(16.0))
    fnp = fn_box.text_frame.paragraphs[0]
    fnp.text = "[1] Phase 2 includes an additional investment resource to support quick win execution and accelerate savings delivery for Telix"
    fnp.font.name = "Aptos"
    fnp.font.size = Pt(7.5)
    fnp.font.italic = True
    fnp.font.color.rgb = hex_to_rgb("64748B")

    # Bottom: GEP Support Infrastructure (top = 408.0, h = 118.0)
    infra_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(408.0), Pt(900.0), Pt(118.0))
    apply_solid_fill(infra_box, "FFFFFF")
    infra_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(infra_box, 4.0)
    apply_library_shadow(infra_box)

    # Infrastructure Header pill
    ih_elem = clean_element(copy.deepcopy(hdr_proto._element))
    slide.shapes._spTree.append(ih_elem)
    sh_ih = slide.shapes[-1]
    sh_ih.left = Pt(360.0)
    sh_ih.top = Pt(398.0)
    sh_ih.width = Pt(240.0)
    sh_ih.height = Pt(22.0)'''

new_s12 = '''    # Bottom: GEP Support Infrastructure (top = 406.0, h = 114.0)
    infra_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(30.0), Pt(406.0), Pt(900.0), Pt(114.0))
    apply_solid_fill(infra_box, "FFFFFF")
    infra_box.line.color.rgb = hex_to_rgb("CBD5E1")
    set_shape_corner_radius(infra_box, 4.0)
    apply_library_shadow(infra_box)

    # Infrastructure Header pill (inside container top edge)
    sh_ih = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(38.0), Pt(412.0), Pt(884.0), Pt(22.0))
    apply_solid_fill(sh_ih, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(sh_ih, 4.0)
    sh_ih.text_frame.word_wrap = False
    sh_ih.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    # Footnote at slide base
    fn_box = slide.shapes.add_textbox(Pt(30.0), Pt(522.0), Pt(900.0), Pt(16.0))
    fnp = fn_box.text_frame.paragraphs[0]
    fnp.text = "[1] Phase 2 includes an additional investment resource to support quick win execution and accelerate savings delivery for Telix"
    fnp.font.name = "Aptos"
    fnp.font.size = Pt(9.0)
    fnp.font.italic = True
    fnp.font.color.rgb = hex_to_rgb("64748B")'''

content = content.replace(old_s12, new_s12)

# Also adjust cards inside infra_box on Slide 12
content = content.replace(
    'ibox = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(ix), Pt(426.0), Pt(iw), Pt(92.0))',
    'ibox = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(ix), Pt(438.0), Pt(iw), Pt(76.0))'
)
content = content.replace(
    'itf_box = slide.shapes.add_textbox(Pt(ix + 4.0), Pt(432.0), Pt(iw - 8.0), Pt(80.0))',
    'itf_box = slide.shapes.add_textbox(Pt(ix + 4.0), Pt(440.0), Pt(iw - 8.0), Pt(72.0))'
)

# 7. Upgrade all .font.size = Pt(< 9.0) to Pt(9.0)
content = re.sub(r'(\.font\.size\s*=\s*)Pt\(([0-8]\.[0-9]+|[0-8])\)', r'\g<1>Pt(9.0)', content)

target.write_text(content, encoding="utf-8")
print("Updated builders_8_to_20.py successfully!")

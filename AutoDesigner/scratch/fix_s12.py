from pathlib import Path

target = Path(r"c:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\scratch\builders_8_to_20.py")
c = target.read_text(encoding="utf-8")

old_block = '''    # Infrastructure Header pill (inside container top edge)
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
    fnp.font.color.rgb = hex_to_rgb("64748B")
    apply_solid_fill(sh_ih, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(sh_ih, 4.0)
    sh_ih.text_frame.word_wrap = False
    sh_ih.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_ih = sh_ih.text_frame.paragraphs[0]
    p_ih.text = "GEP Support Infrastructure"
    p_ih.alignment = PP_ALIGN.CENTER
    p_ih.font.name = "Aptos"
    p_ih.font.size = Pt(10.0)
    p_ih.font.bold = True
    p_ih.font.color.rgb = hex_to_rgb("FFFFFF")

    infra_items = [
        "Global Industry SMEs",
        "Digital Tools & Contracting RPAs",
        "Training Playbooks",
        "Industry Benchmarks",
        "Contracting Process Tools & Templates"
    ]
    card_w = 168.0
    for idx, item_txt in enumerate(infra_items):
        ix = 42.0 + idx * 174.0
        ibox = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(ix), Pt(426.0), Pt(card_w), Pt(90.0))
        apply_solid_fill(ibox, "F8FAFC")
        ibox.line.color.rgb = hex_to_rgb("E2E8F0")
        set_shape_corner_radius(ibox, 4.0)
        apply_library_shadow(ibox)

        # Icon Circle
        ic = slide.shapes.add_shape(MSO_SHAPE.OVAL, Pt(ix + (card_w - 28.0) / 2.0), Pt(434.0), Pt(28.0), Pt(28.0))
        apply_solid_fill(ic, "EFF6FF")
        ic.line.color.rgb = hex_to_rgb("BFDBFE")
        ic_tf = ic.text_frame
        ic_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        icp = ic_tf.paragraphs[0]
        icp.text = "⚙"
        icp.font.size = Pt(12.0)
        icp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)
        icp.alignment = PP_ALIGN.CENTER

        itf_box = slide.shapes.add_textbox(Pt(ix + 4.0), Pt(464.0), Pt(card_w - 8.0), Pt(48.0))'''

new_block = '''    # Infrastructure Header pill (inside container top edge)
    sh_ih = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(38.0), Pt(410.0), Pt(884.0), Pt(20.0))
    apply_solid_fill(sh_ih, ColorTokens.BRAND_NAVY)
    set_shape_corner_radius(sh_ih, 4.0)
    sh_ih.text_frame.word_wrap = False
    sh_ih.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_ih = sh_ih.text_frame.paragraphs[0]
    p_ih.text = "GEP Support Infrastructure"
    p_ih.alignment = PP_ALIGN.CENTER
    p_ih.font.name = "Aptos"
    p_ih.font.size = Pt(10.0)
    p_ih.font.bold = True
    p_ih.font.color.rgb = hex_to_rgb("FFFFFF")

    # Footnote at slide base
    fn_box = slide.shapes.add_textbox(Pt(30.0), Pt(522.0), Pt(900.0), Pt(16.0))
    fnp = fn_box.text_frame.paragraphs[0]
    fnp.text = "[1] Phase 2 includes an additional investment resource to support quick win execution and accelerate savings delivery for Telix"
    fnp.font.name = "Aptos"
    fnp.font.size = Pt(9.0)
    fnp.font.italic = True
    fnp.font.color.rgb = hex_to_rgb("64748B")

    infra_items = [
        "Global Industry SMEs",
        "Digital Tools & Contracting RPAs",
        "Training Playbooks",
        "Industry Benchmarks",
        "Contracting Process Tools & Templates"
    ]
    card_w = 168.0
    for idx, item_txt in enumerate(infra_items):
        ix = 42.0 + idx * 174.0
        ibox = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(ix), Pt(434.0), Pt(card_w), Pt(80.0))
        apply_solid_fill(ibox, "F8FAFC")
        ibox.line.color.rgb = hex_to_rgb("E2E8F0")
        set_shape_corner_radius(ibox, 4.0)
        apply_library_shadow(ibox)

        # Icon Circle
        ic = slide.shapes.add_shape(MSO_SHAPE.OVAL, Pt(ix + (card_w - 24.0) / 2.0), Pt(438.0), Pt(24.0), Pt(24.0))
        apply_solid_fill(ic, "EFF6FF")
        ic.line.color.rgb = hex_to_rgb("BFDBFE")
        ic_tf = ic.text_frame
        ic_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        icp = ic_tf.paragraphs[0]
        icp.text = "⚙"
        icp.font.size = Pt(11.0)
        icp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO)
        icp.alignment = PP_ALIGN.CENTER

        itf_box = slide.shapes.add_textbox(Pt(ix + 4.0), Pt(464.0), Pt(card_w - 8.0), Pt(48.0))'''

assert old_block in c, "old_block not found in c!"
c = c.replace(old_block, new_block)
target.write_text(c, encoding="utf-8")
print("Slide 12 layout successfully updated!")

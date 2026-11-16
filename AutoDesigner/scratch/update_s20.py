from pathlib import Path

target = Path(r"c:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\scratch\builders_8_to_20.py")
content = target.read_text(encoding="utf-8")

old_s20_table = '''        # KPOD Stakeholders (Left)
        kb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(40.0), Pt(c_y + 20.0), Pt(120.0), Pt(46.0))
        apply_solid_fill(kb, "F8FAFC")
        kb.line.color.rgb = hex_to_rgb("CBD5E1")
        set_shape_corner_radius(kb, 4.0)
        ktf = kb.text_frame
        ktf.word_wrap = True
        ktf.vertical_anchor = MSO_ANCHOR.MIDDLE
        kp = ktf.paragraphs[0]
        kp.text = com["kpod"]
        kp.alignment = PP_ALIGN.CENTER
        kp.font.name = "Aptos"
        kp.font.size = Pt(9.0)
        kp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

        # Committee Topic Blocks (Center)
        n_blks = len(com["blocks"])
        blk_w = 350.0 / n_blks
        for b_idx, blk_name in enumerate(com["blocks"]):
            bx = 168.0 + b_idx * blk_w
            bbox = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(bx), Pt(c_y + 20.0), Pt(blk_w - 4.0), Pt(46.0))
            apply_solid_fill(bbox, "EFF6FF" if c_i % 2 == 0 else "FFF7ED")
            bbox.line.color.rgb = hex_to_rgb("BFDBFE" if c_i % 2 == 0 else "FED7AA")
            set_shape_corner_radius(bbox, 4.0)
            btf = bbox.text_frame
            btf.word_wrap = True
            btf.vertical_anchor = MSO_ANCHOR.MIDDLE
            bp = btf.paragraphs[0]
            bp.text = blk_name
            bp.alignment = PP_ALIGN.CENTER
            bp.font.name = "Aptos"
            bp.font.size = Pt(9.0)
            bp.font.bold = True
            bp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO if c_i % 2 == 0 else ColorTokens.BRAND_TANGERINE)

        # GEP Stakeholders (Right)
        gb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(526.0), Pt(c_y + 20.0), Pt(134.0), Pt(46.0))'''

new_s20_table = '''        # KPOD Stakeholders (Left)
        kb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(38.0), Pt(c_y + 20.0), Pt(100.0), Pt(46.0))
        apply_solid_fill(kb, "F8FAFC")
        kb.line.color.rgb = hex_to_rgb("CBD5E1")
        set_shape_corner_radius(kb, 4.0)
        ktf = kb.text_frame
        ktf.word_wrap = True
        ktf.vertical_anchor = MSO_ANCHOR.MIDDLE
        kp = ktf.paragraphs[0]
        kp.text = com["kpod"]
        kp.alignment = PP_ALIGN.CENTER
        kp.font.name = "Aptos"
        kp.font.size = Pt(9.0)
        kp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_NAVY)

        # Committee Topic Blocks (Center)
        n_blks = len(com["blocks"])
        blk_w = 414.0 / n_blks
        for b_idx, blk_name in enumerate(com["blocks"]):
            bx = 142.0 + b_idx * blk_w
            bbox = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(bx), Pt(c_y + 20.0), Pt(blk_w - 4.0), Pt(46.0))
            apply_solid_fill(bbox, "EFF6FF" if c_i % 2 == 0 else "FFF7ED")
            bbox.line.color.rgb = hex_to_rgb("BFDBFE" if c_i % 2 == 0 else "FED7AA")
            set_shape_corner_radius(bbox, 4.0)
            btf = bbox.text_frame
            btf.word_wrap = True
            btf.vertical_anchor = MSO_ANCHOR.MIDDLE
            bp = btf.paragraphs[0]
            bp.text = blk_name
            bp.alignment = PP_ALIGN.CENTER
            bp.font.name = "Aptos"
            bp.font.size = Pt(9.0)
            bp.font.bold = True
            bp.font.color.rgb = hex_to_rgb(ColorTokens.BRAND_INDIGO if c_i % 2 == 0 else ColorTokens.BRAND_TANGERINE)

        # GEP Stakeholders (Right)
        gb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(560.0), Pt(c_y + 20.0), Pt(100.0), Pt(46.0))'''

assert old_s20_table in content, "old_s20_table not found!"
content = content.replace(old_s20_table, new_s20_table)

old_badge = '''    # Left Vertical Header Badge
    avb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(38.0), Pt(348.0), Pt(36.0), Pt(166.0))
    apply_solid_fill(avb, ColorTokens.BRAND_TANGERINE)
    set_shape_corner_radius(avb, 4.0)
    avbtf = avb.text_frame
    avbtf.vertical_anchor = MSO_ANCHOR.MIDDLE
    avbp = avbtf.paragraphs[0]
    avbp.text = "Accountability"
    avbp.alignment = PP_ALIGN.CENTER
    avbp.font.name = "Aptos"
    avbp.font.size = Pt(9.0)
    avbp.font.bold = True
    avbp.font.color.rgb = hex_to_rgb("FFFFFF")'''

new_badge = '''    # Left Vertical Header Badge (Rotated 270 deg vertical text)
    avb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(38.0), Pt(348.0), Pt(36.0), Pt(166.0))
    apply_solid_fill(avb, ColorTokens.BRAND_TANGERINE)
    set_shape_corner_radius(avb, 4.0)
    avbtf = avb.text_frame
    avbtf.word_wrap = False
    avbtf.vertical_anchor = MSO_ANCHOR.MIDDLE
    avbtf._txBody.bodyPr.set("vert", "vert270")
    avbp = avbtf.paragraphs[0]
    avbp.text = "Accountability"
    avbp.alignment = PP_ALIGN.CENTER
    avbp.font.name = "Aptos"
    avbp.font.size = Pt(11.0)
    avbp.font.bold = True
    avbp.font.color.rgb = hex_to_rgb("FFFFFF")'''

assert old_badge in content, "old_badge not found!"
content = content.replace(old_badge, new_badge)

target.write_text(content, encoding="utf-8")
print("Slide 20 columns and vertical badge updated successfully!")

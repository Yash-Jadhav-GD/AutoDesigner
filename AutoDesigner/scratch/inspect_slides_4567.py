import pptx
from pptx.enum.shapes import MSO_SHAPE_TYPE

prs = pptx.Presentation("AutoDesigner/Design_tester_new_inspect.pptx")

for s_num in [4, 5, 6, 7]:
    s = prs.slides[s_num - 1]
    print(f"\n============================= SLIDE {s_num} =============================")
    for idx, sh in enumerate(s.shapes):
        pos = f"x={sh.left.pt:.1f}, y={sh.top.pt:.1f}, w={sh.width.pt:.1f}, h={sh.height.pt:.1f}"
        fill_info = "noFill"
        try:
            if sh.fill.type:
                fill_info = f"fill={sh.fill.type}"
                if sh.fill.type == 1: # solid
                    fill_info += f"({sh.fill.fore_color.rgb})"
        except:
            pass
        
        txts = []
        if sh.has_text_frame:
            for p_i, p in enumerate(sh.text_frame.paragraphs):
                ptxt = p.text.strip()
                if ptxt:
                    sz = p.runs[0].font.size.pt if (p.runs and p.runs[0].font.size) else (p.font.size.pt if p.font.size else "def")
                    b = "B" if (p.runs and p.runs[0].font.bold) or p.font.bold else ""
                    c = ""
                    try:
                        if p.runs and p.runs[0].font.color and p.runs[0].font.color.rgb:
                            c = f"#{p.runs[0].font.color.rgb}"
                    except:
                        pass
                    txts.append(f"p{p_i}(sz={sz},{b},{c}): '{ptxt}'")
        elif sh.has_table:
            t = sh.table
            txts.append(f"TABLE {len(t.rows)}x{len(t.columns)}")
            for r_i, r in enumerate(t.rows):
                for c_i, c in enumerate(r.cells):
                    ct = " ".join(p.text.strip() for p in c.text_frame.paragraphs if p.text.strip())
                    if ct:
                        txts.append(f"  [{r_i},{c_i}]: '{ct}'")
        elif sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            txts.append(f"GROUP with {len(sh.shapes)} subshapes:")
            for g_i, g in enumerate(sh.shapes):
                g_txts = [p.text.strip() for p in g.text_frame.paragraphs if p.text.strip()] if g.has_text_frame else []
                if g_txts:
                    txts.append(f"   sub{g_i} [{g.name}]: {' | '.join(g_txts)}")
        
        print(f"Shape {idx+1} [{sh.name}] ({sh.shape_type}) {pos} {fill_info}")
        for t in txts:
            print(f"    {t}")

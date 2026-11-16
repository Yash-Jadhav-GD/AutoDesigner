import pptx
from pptx.enum.shapes import MSO_SHAPE_TYPE

prs = pptx.Presentation("AutoDesigner/Design_tester_new_inspect.pptx")

for s_num in [2, 3]:
    s = prs.slides[s_num - 1]
    print(f"\n============================= SLIDE {s_num} =============================")
    for idx, sh in enumerate(s.shapes):
        pos = f"x={sh.left.pt:.1f}, y={sh.top.pt:.1f}, w={sh.width.pt:.1f}, h={sh.height.pt:.1f}"
        txts = []
        if sh.has_text_frame:
            for p_i, p in enumerate(sh.text_frame.paragraphs):
                ptxt = p.text.strip()
                if ptxt:
                    sz = p.runs[0].font.size.pt if (p.runs and p.runs[0].font.size) else (p.font.size.pt if p.font.size else "def")
                    b = "B" if (p.runs and p.runs[0].font.bold) or p.font.bold else ""
                    txts.append(f"p{p_i}(sz={sz},{b}): '{ptxt}'")
        elif sh.has_table:
            t = sh.table
            txts.append(f"TABLE {len(t.rows)}x{len(t.columns)}")
        elif sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            txts.append(f"GROUP with {len(sh.shapes)} subshapes")
        
        if txts or sh.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.TABLE):
            print(f"Shape {idx+1} [{sh.name}] ({sh.shape_type}) {pos}")
            for t in txts:
                print(f"    {t}")

import pptx

prs = pptx.Presentation("AutoDesigner/Design_tester_new_inspect.pptx")

for i in range(len(prs.slides)):
    s = prs.slides[i]
    print(f"\n{'='*30} SLIDE {i+1} {'='*30}")
    for idx, sh in enumerate(s.shapes):
        t_list = []
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                if p.text.strip():
                    t_list.append(p.text.strip())
        elif sh.has_table:
            for r in sh.table.rows:
                for c in r.cells:
                    ct = " ".join(p.text.strip() for p in c.text_frame.paragraphs if p.text.strip())
                    if ct:
                        t_list.append(ct)
        elif sh.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.GROUP:
            for g_sh in sh.shapes:
                if g_sh.has_text_frame:
                    for p in g_sh.text_frame.paragraphs:
                        if p.text.strip():
                            t_list.append(f"[G:{p.text.strip()}]")
        
        if t_list or sh.shape_type in (pptx.enum.shapes.MSO_SHAPE_TYPE.PICTURE, pptx.enum.shapes.MSO_SHAPE_TYPE.TABLE):
            print(f"Shape {idx+1} [{sh.name}] ({sh.shape_type}):")
            for t in t_list:
                print(f"    - {t}")

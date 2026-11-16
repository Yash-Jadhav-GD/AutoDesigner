import pptx

prs = pptx.Presentation("AutoDesigner/Design_tester_new_inspect.pptx")

for i in range(len(prs.slides)):
    s = prs.slides[i]
    all_texts = []
    for sh in s.shapes:
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                t = p.text.strip()
                if t and t != ".":
                    all_texts.append(t)
        elif sh.has_table:
            for r in sh.table.rows:
                for c in r.cells:
                    for p in c.text_frame.paragraphs:
                        t = p.text.strip()
                        if t:
                            all_texts.append(t)
        elif sh.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.GROUP:
            for g in sh.shapes:
                if g.has_text_frame:
                    for p in g.text_frame.paragraphs:
                        t = p.text.strip()
                        if t and t != ".":
                            all_texts.append(t)
    print(f"Slide {i+1}: {len(all_texts)} text items")

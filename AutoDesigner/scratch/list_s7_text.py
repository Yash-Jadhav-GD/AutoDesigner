import pptx

prs = pptx.Presentation("AutoDesigner/Design_tester_new_inspect.pptx")
s7 = prs.slides[6]
items = []
for sh in s7.shapes:
    if sh.has_text_frame:
        for p in sh.text_frame.paragraphs:
            if p.text.strip() and p.text.strip() != '.':
                items.append((sh.name, p.text.strip()))
    elif sh.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.GROUP:
        for g in sh.shapes:
            if g.has_text_frame:
                for p in g.text_frame.paragraphs:
                    if p.text.strip() and p.text.strip() != '.':
                        items.append((f"{sh.name}->{g.name}", p.text.strip()))

print(f"Total valid text items on Slide 7: {len(items)}")
for idx, (src, txt) in enumerate(items):
    print(f"{idx+1:2d}. [{src}]: \"{txt}\"")

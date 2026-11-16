import pptx

prs = pptx.Presentation("AutoDesigner/Design_tester_new_inspect.pptx")
s3 = prs.slides[2]
items = []
for sh in s3.shapes:
    if sh.has_text_frame:
        for p in sh.text_frame.paragraphs:
            if p.text.strip():
                items.append((sh.name, p.text.strip()))

print(f"Total text items on Slide 3: {len(items)}")
for idx, (src, txt) in enumerate(items):
    print(f"{idx+1:2d}. [{src}]: \"{txt[:80]}\"")

import pptx

prs = pptx.Presentation("AutoDesigner/Library_copy.pptx")
print(f"Total slides: {len(prs.slides)}")
for i, s in enumerate(prs.slides, 1):
    names = [sh.name for sh in s.shapes]
    print(f"Slide {i}: {len(s.shapes)} shapes: {names[:6]}")

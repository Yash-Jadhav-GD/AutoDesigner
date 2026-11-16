import pptx
from pathlib import Path

prs = pptx.Presentation("AutoDesigner/Design_tester_new_inspect.pptx")

for i, s in enumerate(prs.slides):
    pics = [sh for sh in s.shapes if sh.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.PICTURE]
    if pics:
        print(f"Slide {i+1} has {len(pics)} picture(s):")
        for p in pics:
            print(f"  [{p.name}] format={p.image.content_type}, size={p.image.size}, w={p.width.pt:.1f}pt, h={p.height.pt:.1f}pt")

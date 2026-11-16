from pptx import Presentation

prs = Presentation("AutoDesigner/Design tester new.pptx")
for i in range(7, len(prs.slides)):
    s = prs.slides[i]
    print(f"\n=== SLIDE {i+1} Pictures / Media ===")
    for sh in s.shapes:
        if sh.shape_type == 13: # MSO_SHAPE_TYPE.PICTURE
            print(f"  Picture '{sh.name}' at ({round(sh.left.pt,1)}, {round(sh.top.pt,1)}, w={round(sh.width.pt,1)}, h={round(sh.height.pt,1)})")
        elif sh.shape_type == 6: # GROUP
            for sub in sh.shapes:
                if sub.shape_type == 13:
                    print(f"  Grouped Picture '{sub.name}' in '{sh.name}' at ({round(sub.left.pt,1)}, {round(sub.top.pt,1)})")

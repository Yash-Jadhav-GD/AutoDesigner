from pptx import Presentation
import io

prs = Presentation("AutoDesigner/Design tester new.pptx")
for i in [8, 11, 16]: # slides 9, 12, 17
    s = prs.slides[i]
    print(f"Slide {i+1} has {len(s.shapes)} shapes")
    for sh in s.shapes:
        if sh.shape_type == 13: # PICTURE
            print(f"  Picture {sh.name}: {len(sh.image.blob)} bytes, ext: {sh.image.ext}")
        elif sh.shape_type == 6: # GROUP
            for sub in sh.shapes:
                if sub.shape_type == 13:
                    print(f"  Grouped Picture {sub.name}: {len(sub.image.blob)} bytes, ext: {sub.image.ext}")

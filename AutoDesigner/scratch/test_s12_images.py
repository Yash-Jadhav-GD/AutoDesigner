from pptx import Presentation

prs = Presentation("AutoDesigner/Design tester new.pptx")
s12 = prs.slides[11]
for sh in s12.shapes:
    if sh.shape_type == 13:
        try:
            blob = sh.image.blob
            print(f"Picture {sh.name}: {len(blob)} bytes, ext: {sh.image.ext}")
        except Exception as e:
            print(f"Picture {sh.name} error: {e}")

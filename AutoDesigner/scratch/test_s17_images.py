from pptx import Presentation
import io

prs = Presentation("AutoDesigner/Design tester new.pptx")
s17 = prs.slides[16] # 0-indexed 16 is slide 17
for sh in s17.shapes:
    if sh.shape_type == 13:
        try:
            blob = sh.image.blob
            print(f"Picture {sh.name}: {len(blob)} bytes, ext: {sh.image.ext}")
        except Exception as e:
            print(f"Picture {sh.name} error: {e}")

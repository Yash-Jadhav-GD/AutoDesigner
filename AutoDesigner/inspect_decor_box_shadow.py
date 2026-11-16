import pptx
from xml.etree import ElementTree as ET

prs = pptx.Presentation("AutoDesigner/Library_copy.pptx")
print(f"Total slides in Library_copy.pptx: {len(prs.slides)}")
for s_idx, slide in enumerate(prs.slides, 1):
    for sh in slide.shapes:
        if sh.name == "Decor_Box_with_Shadow_1":
            print(f"Found '{sh.name}' on Slide {s_idx}!")
            print(f"Type: {sh.shape_type}")
            print(f"Position: left={sh.left.pt:.1f}pt, top={sh.top.pt:.1f}pt, width={sh.width.pt:.1f}pt, height={sh.height.pt:.1f}pt")
            print("Full XML:\n", ET.tostring(sh._element, encoding='utf-8').decode('utf-8'))

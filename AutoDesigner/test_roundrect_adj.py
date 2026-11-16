import pptx
from pptx.util import Pt
from pptx.oxml import parse_xml

prs = pptx.Presentation()
prs.slide_width = Pt(960)
prs.slide_height = Pt(540)
slide = prs.slides.add_slide(prs.slide_layouts[6])

# Add a rounded rectangle
sh = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(100), Pt(100), Pt(300), Pt(200))
print("Default adjustments:")
print("sh.adjustments count:", len(sh.adjustments))
for i in range(len(sh.adjustments)):
    print(f"adj[{i}] = {sh.adjustments[i]}")

# Look at the raw XML of the shape
from xml.etree import ElementTree as ET
xml_str = ET.tostring(sh._element, encoding='utf-8').decode('utf-8')
print("Shape XML:\n", xml_str[:400])

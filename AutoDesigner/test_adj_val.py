import pptx
from pptx.util import Pt
from xml.etree import ElementTree as ET

prs = pptx.Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
sh = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(100), Pt(100), Pt(300), Pt(200))
print("Before setting adj:")
print(ET.tostring(sh._element.spPr.prstGeom, encoding='utf-8').decode('utf-8'))

sh.adjustments[0] = 0.05
print("\nAfter setting adj to 0.05:")
print(ET.tostring(sh._element.spPr.prstGeom, encoding='utf-8').decode('utf-8'))
print("adj[0] is now:", sh.adjustments[0])

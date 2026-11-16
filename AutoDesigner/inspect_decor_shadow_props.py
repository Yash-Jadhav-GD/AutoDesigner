import pptx
from xml.etree import ElementTree as ET

prs = pptx.Presentation("AutoDesigner/Library_copy.pptx")
sh = [s for s in prs.slides[5].shapes if s.name == "Decor_Box_with_Shadow_1"][0]

print("Shape name:", sh.name)
print("Shape type:", sh.shape_type)
spPr = sh._element.spPr
print("spPr tags:", [c.tag.split("}")[-1] for c in spPr])

effectLst = spPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}effectLst")
print("effectLst XML:\n", ET.tostring(effectLst, encoding='utf-8').decode('utf-8'))

ln = spPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}ln")
if ln is not None:
    print("ln XML:\n", ET.tostring(ln, encoding='utf-8').decode('utf-8'))

fill = spPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill")
if fill is not None:
    print("solidFill XML:\n", ET.tostring(fill, encoding='utf-8').decode('utf-8'))

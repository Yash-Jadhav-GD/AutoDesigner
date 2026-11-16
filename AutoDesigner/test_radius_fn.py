import pptx
from pptx.util import Pt
from xml.etree import ElementTree as ET

def set_shape_corner_radius(shape, radius_pt=4.0):
    try:
        prst = shape._element.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom")
        if prst is not None and "round" in prst.get("prst", ""):
            w = shape.width.pt
            h = shape.height.pt
            min_dim = min(w, h)
            if min_dim > 0:
                shape.adjustments[0] = min(0.5, max(0.0, radius_pt / min_dim))
                return True
    except Exception as e:
        print("Error setting radius:", e)
    return False

prs = pptx.Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
s1 = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(10), Pt(10), Pt(100), Pt(200))
s2 = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.RECTANGLE, Pt(120), Pt(10), Pt(100), Pt(200))
s3 = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.OVAL, Pt(230), Pt(10), Pt(100), Pt(100))

print("s1 rounded rect:", set_shape_corner_radius(s1, 4.0), "adj:", s1.adjustments[0])
print("s2 rect:", set_shape_corner_radius(s2, 4.0))
print("s3 oval:", set_shape_corner_radius(s3, 4.0))

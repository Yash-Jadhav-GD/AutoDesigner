import pptx
from pptx.util import Pt
from pptx.dml.color import RGBColor
import io

prs = pptx.Presentation()
prs.slide_width = Pt(960)
prs.slide_height = Pt(540)
blank = prs.slide_layouts[6]
slide = prs.slides.add_slide(blank)

def add_connector_line(slide, x1, y1, x2, y2, color_hex="A2ABB0", width_pt=1.0):
    conn = slide.shapes.add_connector(pptx.enum.shapes.MSO_CONNECTOR.STRAIGHT, Pt(x1), Pt(y1), Pt(x2), Pt(y2))
    r = int(color_hex[0:2], 16)
    g = int(color_hex[2:4], 16)
    b = int(color_hex[4:6], 16)
    conn.line.color.rgb = RGBColor(r, g, b)
    conn.line.width = Pt(width_pt)
    return conn

add_connector_line(slide, 480, 50, 480, 100)
add_connector_line(slide, 100, 100, 800, 100)

buf = io.BytesIO()
prs.save(buf)
print("Connector test passed! Bytes:", len(buf.getvalue()))

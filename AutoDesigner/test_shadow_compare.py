import pptx
from pptx.util import Pt
from pptx.oxml import parse_xml

prs = pptx.Presentation()
prs.slide_width = Pt(960)
prs.slide_height = Pt(540)
s = prs.slides.add_slide(prs.slide_layouts[6])

def apply_shdw(sh):
    eff_xml = (
        '<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:outerShdw blurRad="190500" rotWithShape="0">'
        '<a:srgbClr val="95939D">'
        '<a:alpha val="16000"/>'
        '</a:srgbClr>'
        '</a:outerShdw>'
        '</a:effectLst>'
    )
    sh._element.spPr.append(parse_xml(eff_xml))

# Card 1: Only outer card has shadow
c1 = s.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(50), Pt(50), Pt(200), Pt(300))
c1.fill.solid()
c1.fill.fore_color.rgb = pptx.dml.color.RGBColor(255, 255, 255)
c1.adjustments[0] = 4.0 / 200.0
apply_shdw(c1)

t1 = s.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(50), Pt(50), Pt(200), Pt(30))
t1.fill.solid()
t1.fill.fore_color.rgb = pptx.dml.color.RGBColor(0x48, 0x39, 0x7E)
t1.line.fill.background()
t1.adjustments[0] = 4.0 / 30.0

# Card 2: Both card and tab have shadow
c2 = s.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(300), Pt(50), Pt(200), Pt(300))
c2.fill.solid()
c2.fill.fore_color.rgb = pptx.dml.color.RGBColor(255, 255, 255)
c2.adjustments[0] = 4.0 / 200.0
apply_shdw(c2)

t2 = s.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(300), Pt(50), Pt(200), Pt(30))
t2.fill.solid()
t2.fill.fore_color.rgb = pptx.dml.color.RGBColor(0x48, 0x39, 0x7E)
t2.line.fill.background()
t2.adjustments[0] = 4.0 / 30.0
apply_shdw(t2)

prs.save("AutoDesigner/test_shadow_compare.pptx")
print("Saved test_shadow_compare.pptx")

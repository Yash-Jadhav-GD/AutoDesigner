import pptx
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.oxml import parse_xml
import io

prs = pptx.Presentation()
prs.slide_width = Pt(960)
prs.slide_height = Pt(540)
blank = prs.slide_layouts[6]
slide = prs.slides.add_slide(blank)

shape = slide.shapes.add_table(2, 2, Pt(30), Pt(60), Pt(400), Pt(100))
tbl = shape.table

def set_cell_border(cell, color="E2E8F0", width="6350"):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.set("anchor", "ctr")
    for border_name in ["lnL", "lnR", "lnT", "lnB"]:
        existing = tcPr.find(f"{{http://schemas.openxmlformats.org/drawingml/2006/main}}{border_name}")
        if existing is not None:
            tcPr.remove(existing)
        ln_xml = f'<a:{border_name} xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="{width}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:{border_name}>'
        tcPr.append(parse_xml(ln_xml))

set_cell_border(tbl.cell(0, 0))
buf = io.BytesIO()
prs.save(buf)
print("Border styling test passed! Bytes:", len(buf.getvalue()))

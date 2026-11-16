import sys
from pathlib import Path
sys.path.insert(0, "AutoDesigner")
from engine.color_tokens import apply_linear_gradient

import pptx
from pptx.util import Pt
prs = pptx.Presentation()
s = prs.slides.add_slide(prs.slide_layouts[6])
sh = s.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(10), Pt(10), Pt(100), Pt(50))
apply_linear_gradient(sh, "48397E", "1F2E65", with_shadow=True)
from xml.etree import ElementTree as ET
eff = sh._element.spPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}effectLst")
print("apply_linear_gradient shadow:\n", ET.tostring(eff, encoding='utf-8').decode('utf-8'))

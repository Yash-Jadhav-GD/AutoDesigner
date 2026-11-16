import sys
from pathlib import Path
BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))
import pptx
from pptx.util import Pt
from engine.builders.common import load_library, clean_element
import copy
from xml.etree import ElementTree as ET

prs = pptx.Presentation()
prs.slide_width = Pt(960)
prs.slide_height = Pt(540)
slide = prs.slides.add_slide(prs.slide_layouts[6])
lib_prs, lib_s1, lib_s3, lib_lines = load_library()

# Test 1: newly created rounded rectangle
sh1 = slide.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Pt(50), Pt(50), Pt(200), Pt(100))
w = sh1.width.pt
h = sh1.height.pt
sh1.adjustments[0] = 4.0 / min(w, h)
print(f"sh1 adjustments: {sh1.adjustments[0]}")
av = sh1._element.spPr.prstGeom.avLst
print(f"sh1 avLst: {[g.attrib for g in av]}")

# Test 2: cloned from library prototype
d1_box_proto = [s for s in lib_s1.shapes if s.name == "Design 1 Decor box"][0]
elem = clean_element(copy.deepcopy(d1_box_proto._element))
slide.shapes._spTree.append(elem)
sh2 = slide.shapes[len(slide.shapes) - 1]
sh2.left = Pt(300)
sh2.top = Pt(50)
sh2.width = Pt(200)
sh2.height = Pt(100)
print(f"sh2 initial adjustments: {sh2.adjustments[0]}")
sh2.adjustments[0] = 4.0 / min(sh2.width.pt, sh2.height.pt)
print(f"sh2 new adjustments: {sh2.adjustments[0]}")
av2 = sh2._element.spPr.prstGeom.avLst
print(f"sh2 avLst: {[g.attrib for g in av2]}")

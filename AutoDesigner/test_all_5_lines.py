import sys
from pathlib import Path
BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))
import pptx
from pptx.util import Pt
from engine.builders.common import load_library, add_library_line, clean_element
import copy
from xml.etree import ElementTree as ET

prs = pptx.Presentation()
prs.slide_width = Pt(960)
prs.slide_height = Pt(540)
slide = prs.slides.add_slide(prs.slide_layouts[6])
_, _, _, lib_lines = load_library()

def add_lib_line_clean(slide, proto_shape, x_pt, y_pt, length_pt, is_vertical=False):
    elem = clean_element(copy.deepcopy(proto_shape._element))
    xfrm = elem.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm")
    off = xfrm.find("{http://schemas.openxmlformats.org/drawingml/2006/main}off")
    ext = xfrm.find("{http://schemas.openxmlformats.org/drawingml/2006/main}ext")
    
    # Remove rotation attribute if present so positioning is pure and predictable
    if "rot" in xfrm.attrib:
        del xfrm.attrib["rot"]
        
    emu_x = int(x_pt * 12700)
    emu_y = int(y_pt * 12700)
    emu_len = int(length_pt * 12700)
    
    off.set("x", str(emu_x))
    off.set("y", str(emu_y))
    
    if is_vertical:
        ext.set("cx", "0")
        ext.set("cy", str(emu_len))
    else:
        ext.set("cx", str(emu_len))
        ext.set("cy", "0")
        
    slide.shapes._spTree.append(elem)
    return slide.shapes[len(slide.shapes) - 1]

# Test all 5 lines!
add_lib_line_clean(slide, lib_lines["side_fade_r"], 50.0, 50.0, 300.0, is_vertical=False)
add_lib_line_clean(slide, lib_lines["centre_fade_h"], 50.0, 100.0, 300.0, is_vertical=False)
add_lib_line_clean(slide, lib_lines["solid_h"], 50.0, 150.0, 300.0, is_vertical=False)
add_lib_line_clean(slide, lib_lines["solid_v"], 400.0, 50.0, 200.0, is_vertical=True)
add_lib_line_clean(slide, lib_lines["centre_fade_v"], 450.0, 50.0, 200.0, is_vertical=True)

prs.save("AutoDesigner/test_all_5_lines.pptx")
print("Saved test_all_5_lines.pptx!")

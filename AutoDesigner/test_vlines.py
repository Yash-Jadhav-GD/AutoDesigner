import sys, os
from pathlib import Path
BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))
import pptx
from pptx.util import Pt
from engine.builders.common import load_library, add_library_line

prs = pptx.Presentation()
prs.slide_width = Pt(960)
prs.slide_height = Pt(540)
slide = prs.slides.add_slide(prs.slide_layouts[6])

_, _, _, lib_lines = load_library()
print("Loaded lib_lines:", list(lib_lines.keys()))

# Test adding Grey Solid V
line_v = add_library_line(slide, lib_lines["solid_v"], 100.0, 100.0, 0.0, 150.0)
print(f"Added line_v: left={line_v.left.pt}, top={line_v.top.pt}, width={line_v.width.pt}, height={line_v.height.pt}")

# Test adding Grey Centre Fade V
line_cfv = add_library_line(slide, lib_lines["centre_fade_v"], 200.0, 100.0, 0.0, 150.0)
print(f"Added line_cfv: left={line_cfv.left.pt}, top={line_cfv.top.pt}, width={line_cfv.width.pt}, height={line_cfv.height.pt}")

prs.save("AutoDesigner/test_vlines.pptx")
print("Saved test_vlines.pptx successfully!")

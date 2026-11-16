import sys
from pathlib import Path
BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))
import pptx
from pptx.util import Pt
from pptx.oxml import parse_xml
import copy

from engine.builders.common import load_library

lib_prs, lib_s1, lib_s3, lib_lines = load_library()
decor_shadow_proto = None
for s in lib_prs.slides:
    for sh in s.shapes:
        if sh.name == "Decor_Box_with_Shadow_1":
            decor_shadow_proto = sh
            break

print("Found decor_shadow_proto:", decor_shadow_proto is not None)
p_eff = decor_shadow_proto._element.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}effectLst")
print("Proto effectLst tag:", p_eff.tag if p_eff is not None else None)

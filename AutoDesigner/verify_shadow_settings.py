import sys
from pathlib import Path
BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))
import pptx
from xml.etree import ElementTree as ET

prs = pptx.Presentation(BASE_DIR / "output" / "Design_tester_new_formatted.pptx")

for s_idx, slide in enumerate(prs.slides, 1):
    print(f"\n=== Slide {s_idx} ===")
    for sh in slide.shapes:
        sp_xml = sh._element
        prst = sp_xml.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom")
        geom_name = prst.get("prst") if prst is not None else ""
        if "round" in geom_name:
            eff = sp_xml.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}effectLst")
            shdw = eff.find("{http://schemas.openxmlformats.org/drawingml/2006/main}outerShdw") if eff is not None else None
            shdw_desc = "None"
            if shdw is not None:
                clr = shdw.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr")
                clr_val = clr.get("val") if clr is not None else "none"
                alpha = clr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}alpha") if clr is not None else None
                alpha_val = alpha.get("val") if alpha is not None else "none"
                shdw_desc = f"outerShdw: blur={shdw.get('blurRad')}, rot={shdw.get('rotWithShape')}, color=#{clr_val}, alpha={alpha_val}"
            print(f"  Shape '{sh.name}': {shdw_desc}")

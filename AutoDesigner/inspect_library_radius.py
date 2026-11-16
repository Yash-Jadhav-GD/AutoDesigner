import pptx
from pptx.util import Pt
from xml.etree import ElementTree as ET

prs = pptx.Presentation("AutoDesigner/Library_inspect.pptx")
print(f"Slide count: {len(prs.slides)}")

for s_idx, slide in enumerate(prs.slides, 1):
    for sh in slide.shapes:
        sp_xml = sh._element
        prst = sp_xml.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom")
        if prst is not None and prst.get("prst") == "roundRect":
            av = prst.find("{http://schemas.openxmlformats.org/drawingml/2006/main}avLst")
            gd_val = None
            if av is not None:
                gd = av.find("{http://schemas.openxmlformats.org/drawingml/2006/main}gd")
                if gd is not None:
                    gd_val = gd.get("fmla")
            w_pt = sh.width.pt
            h_pt = sh.height.pt
            print(f"Slide {s_idx} shape '{sh.name}': size=({w_pt:.1f}pt, {h_pt:.1f}pt), adj={gd_val}")

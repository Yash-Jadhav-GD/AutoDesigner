import sys
from pathlib import Path
BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))
import pptx
from pptx.util import Pt
from xml.etree import ElementTree as ET

prs = pptx.Presentation(BASE_DIR / "Design_tester_new_inspect.pptx")
print(f"Total slides: {len(prs.slides)}")

s7 = prs.slides[6] # index 6 is Slide 7
print(f"\n==========================================")
print(f"SLIDE 7 SHAPES (USER EDITED SLIDE 6): {len(s7.shapes)}")
print(f"==========================================")

for i, sh in enumerate(s7.shapes, 1):
    sp_xml = sh._element
    tag = sp_xml.tag.split("}")[-1]
    name = sh.name
    pos = f"left={sh.left.pt:.1f}pt, top={sh.top.pt:.1f}pt, width={sh.width.pt:.1f}pt, height={sh.height.pt:.1f}pt"
    
    # Check geometry
    prst = sp_xml.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom")
    geom = prst.get("prst") if prst is not None else "none"
    adj = None
    if prst is not None:
        av = prst.find("{http://schemas.openxmlformats.org/drawingml/2006/main}avLst")
        if av is not None:
            gd = av.find("{http://schemas.openxmlformats.org/drawingml/2006/main}gd")
            if gd is not None:
                adj = gd.get("fmla")

    # Check fill
    solid = sp_xml.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill")
    fill_desc = "noFill"
    if solid is not None:
        clr = solid.find("{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr")
        if clr is not None:
            fill_desc = f"#{clr.get('val')}"
        else:
            sch = solid.find("{http://schemas.openxmlformats.org/drawingml/2006/main}schemeClr")
            fill_desc = f"scheme:{sch.get('val') if sch is not None else '?'}"
            
    # Check line / border
    ln = sp_xml.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}ln")
    ln_desc = "noLine"
    if ln is not None:
        ln_w = ln.get("w")
        ln_clr = ln.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr")
        ln_c = f"#{ln_clr.get('val')}" if ln_clr is not None else "none"
        ln_desc = f"w={ln_w}, clr={ln_c}"

    # Check shadow
    eff = sp_xml.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}effectLst")
    shdw_desc = "noShadow"
    if eff is not None:
        shdw = eff.find("{http://schemas.openxmlformats.org/drawingml/2006/main}outerShdw")
        if shdw is not None:
            shdw_desc = f"outerShdw(blur={shdw.get('blurRad')}, dist={shdw.get('dist')}, dir={shdw.get('dir')})"

    # Check text
    text_summary = []
    if sh.has_text_frame:
        for p in sh.text_frame.paragraphs:
            p_txt = p.text.strip()
            if p_txt:
                font_name = p.font.name if p.font else None
                font_size = p.font.size.pt if p.font and p.font.size else None
                bold = p.font.bold if p.font else None
                runs_info = []
                for r in p.runs:
                    r_sz = r.font.size.pt if r.font and r.font.size else font_size
                    r_b = r.font.bold if r.font else bold
                    r_clr = r.font.color.rgb if r.font and r.font.color and hasattr(r.font.color, 'rgb') else None
                    runs_info.append(f"'{r.text}'(sz={r_sz},b={r_b},c={r_clr})")
                text_summary.append(f"P: {' '.join(runs_info) if runs_info else p_txt}")

    print(f"Shape #{i} [{name}] ({tag}, geom={geom}, adj={adj})")
    print(f"  {pos} | fill={fill_desc} | ln={ln_desc} | {shdw_desc}")
    if text_summary:
        print(f"  Text: {text_summary[:3]}")

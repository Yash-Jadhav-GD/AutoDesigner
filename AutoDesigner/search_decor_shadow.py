import pptx
from pathlib import Path
from xml.etree import ElementTree as ET

files = [
    "Library.pptx",
    "Library_backup.pptx",
    "Library_Modified.pptx",
    "AutoDesigner/Library_copy.pptx",
    "AutoDesigner/Library_inspect.pptx",
    "AutoDesigner/templates/Library.pptx"
]

for f in files:
    try:
        prs = pptx.Presentation(f)
        print(f"\n=== {f} ({len(prs.slides)} slides) ===")
        for s_idx, slide in enumerate(prs.slides, 1):
            for sh in slide.shapes:
                if any(k in sh.name.lower() for k in ["decor", "shadow", "box"]):
                    print(f"  Slide {s_idx} shape: '{sh.name}'")
                    effectLst = sh._element.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}effectLst")
                    if effectLst is not None and len(effectLst) > 0:
                        print(f"    -> effectLst: {ET.tostring(effectLst, encoding='utf-8').decode('utf-8')}")
    except Exception as e:
        print(f"Could not open {f}: {e}")

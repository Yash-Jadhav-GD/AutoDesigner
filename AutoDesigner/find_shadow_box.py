import pptx
from xml.etree import ElementTree as ET

for path in ["AutoDesigner/Library_copy.pptx", "AutoDesigner/Library_inspect.pptx"]:
    try:
        prs = pptx.Presentation(path)
        print(f"=== Checking {path} ===")
        for s_idx, slide in enumerate(prs.slides, 1):
            for sh in slide.shapes:
                if "shadow" in sh.name.lower() or "decor_box" in sh.name.lower() or "decor" in sh.name.lower():
                    print(f"Slide {s_idx} shape: '{sh.name}'")
                    effectLst = sh._element.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}effectLst")
                    if effectLst is not None:
                        print(f"  effectLst XML:\n  {ET.tostring(effectLst, encoding='utf-8').decode('utf-8')}")
                    else:
                        print("  No effectLst found directly in shape")
        break
    except Exception as e:
        print(f"Error checking {path}: {e}")

import zipfile
from xml.etree import ElementTree as ET

for fn in ["AutoDesigner/After_copy.pptx", "AutoDesigner/Library_inspect.pptx"]:
    try:
        with zipfile.ZipFile(fn) as z:
            for name in z.namelist():
                if name.startswith("ppt/slides/slide") and name.endswith(".xml"):
                    root = ET.fromstring(z.read(name))
                    for sp in root.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom"):
                        if "round" in sp.get("prst", ""):
                            av = sp.find("{http://schemas.openxmlformats.org/drawingml/2006/main}avLst")
                            gd = [f"{g.get('name')}={g.get('fmla')}" for g in av.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}gd")] if av is not None else []
                            print(f"{fn} {name}: prst={sp.get('prst')} avLst={gd}")
    except Exception as e:
        print(f"{fn}: {e}")

import sys
from pathlib import Path
BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))
from engine.builders.common import load_library
from xml.etree import ElementTree as ET

_, _, _, lib_lines = load_library()
for k, line in lib_lines.items():
    xfrm = line._element.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm")
    ext = xfrm.find("{http://schemas.openxmlformats.org/drawingml/2006/main}ext") if xfrm is not None else None
    cx = ext.get("cx") if ext is not None else None
    cy = ext.get("cy") if ext is not None else None
    print(f"Key: {k}, name: '{line.name}', cx: {cx}, cy: {cy}")

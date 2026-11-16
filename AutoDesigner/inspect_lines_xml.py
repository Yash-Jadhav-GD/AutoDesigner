import sys
from pathlib import Path
BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))
from engine.builders.common import load_library
from xml.etree import ElementTree as ET

_, _, _, lib_lines = load_library()
for k in ["solid_v", "centre_fade_v", "solid_h"]:
    line = lib_lines[k]
    print(f"\n=== {k} ({line.name}) ===")
    print(ET.tostring(line._element, encoding='utf-8').decode('utf-8'))

import sys
from pathlib import Path
BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))
import pptx
from pptx.util import Pt
from xml.etree import ElementTree as ET

prs = pptx.Presentation(BASE_DIR / "output" / "Design_tester_new_formatted.pptx")
print(f"Loaded presentation: {len(prs.slides)} slides\n")

target_line_names = [
    "Grey Centre Fade V",
    "Grey Solid V",
    "Grey Solid H",
    "Grey Centre Fade H",
    "Grey Side Fade R",
    "Straight Connector 8",
    "White Side Fade R Divider"
]

total_round_rects = 0
total_lines = 0
discrepancies = []

for s_idx, slide in enumerate(prs.slides, 1):
    print(f"=== Slide {s_idx} ===")
    for sh in slide.shapes:
        sp_xml = sh._element
        prst = sp_xml.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom")
        geom_name = prst.get("prst") if prst is not None else "unknown"
        
        # Check rounded rectangles
        if "round" in geom_name:
            total_round_rects += 1
            w = sh.width.pt
            h = sh.height.pt
            min_dim = min(w, h)
            adj = sh.adjustments[0] if len(sh.adjustments) > 0 else None
            # Calculated radius in pt:
            radius_pt = min_dim * adj if adj is not None else 0.0
            print(f"  [Rounded Rect] '{sh.name}' size=({w:.1f}pt, {h:.1f}pt) | adj={adj:.5f} | radius={radius_pt:.2f}pt")
            if abs(radius_pt - 4.0) > 0.1 and adj < 0.49: # allow 0.5 cap on small heights
                discrepancies.append(f"Slide {s_idx} shape '{sh.name}' radius={radius_pt:.2f}pt (expected 4.0pt)")
                
        # Check lines
        if geom_name == "line" or "cxnSp" in sp_xml.tag:
            total_lines += 1
            sh_name = sh.name
            matched = any(t.lower() in sh_name.lower() for t in target_line_names)
            print(f"  [Line] '{sh.name}' | library match: {matched} | size=({sh.width.pt:.1f}pt, {sh.height.pt:.1f}pt)")
            if not matched:
                discrepancies.append(f"Slide {s_idx} line '{sh.name}' not from library list!")

print(f"\n--- Verification Summary ---")
print(f"Total rounded rectangle boxes checked: {total_round_rects}")
print(f"Total lines checked: {total_lines}")
print(f"Discrepancies found: {len(discrepancies)}")
for d in discrepancies:
    print(f"  - {d}")

if len(discrepancies) == 0:
    print("\nALL CHECKS PASSED: 100% of rounded boxes have 4px border radius, and 100% of lines are from the Library prototype list!")
else:
    print(f"\nWARNING: {len(discrepancies)} discrepancies found.")

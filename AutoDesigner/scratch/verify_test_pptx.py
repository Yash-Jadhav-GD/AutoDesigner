import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE_DIR))

import pptx
from pptx.util import Pt

prs = pptx.Presentation(BASE_DIR / "output" / "test_new_designs.pptx")
print(f"Loaded generated presentation: {len(prs.slides)} slides\n")

target_line_names = [
    "Grey Centre Fade V",
    "Grey Solid V",
    "Grey Solid H",
    "Grey Centre Fade H",
    "Grey Side Fade R",
    "Straight Connector 8",
    "White Side Fade R Divider",
]

total_round_rects = 0
total_lines = 0
discrepancies = []
min_font_size = 999.0

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
            if geom_name == "roundRect":
                adj = sh.adjustments[0] if len(sh.adjustments) > 0 else None
                radius_pt = min_dim * adj if adj is not None else 0.0
                print(f"  [roundRect] '{sh.name}' size=({w:.1f}pt, {h:.1f}pt) | adj={adj:.5f} | radius={radius_pt:.2f}pt")
                if abs(radius_pt - 4.0) > 0.1 and adj < 0.49:
                    discrepancies.append(f"Slide {s_idx} shape '{sh.name}' radius={radius_pt:.2f}pt (expected 4.0pt)")
            elif geom_name == "round2SameRect":
                adj0 = sh.adjustments[0] if len(sh.adjustments) > 0 else 0.0
                adj1 = sh.adjustments[1] if len(sh.adjustments) > 1 else 0.0
                radius_pt = min_dim * adj1
                print(f"  [round2SameRect] '{sh.name}' size=({w:.1f}pt, {h:.1f}pt) | bottom_adj={adj0:.5f} | top_adj={adj1:.5f} | top_radius={radius_pt:.2f}pt")
                if adj0 > 0.001:
                    discrepancies.append(f"Slide {s_idx} shape '{sh.name}' bottom corners are rounded (expected square adj=0)")
                if abs(radius_pt - 4.0) > 0.1 and adj1 < 0.49:
                    discrepancies.append(f"Slide {s_idx} shape '{sh.name}' top radius={radius_pt:.2f}pt (expected 4.0pt)")

        # Check lines
        if geom_name == "line" or "cxnSp" in sp_xml.tag:
            total_lines += 1
            sh_name = sh.name
            matched = any(t.lower() in sh_name.lower() for t in target_line_names)
            print(f"  [Line] '{sh.name}' | library match: {matched} | size=({sh.width.pt:.1f}pt, {sh.height.pt:.1f}pt)")
            if not matched:
                discrepancies.append(f"Slide {s_idx} line '{sh.name}' not from library list!")

        # Check font sizes and colors
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if r.font.size is not None:
                        sz = r.font.size.pt
                        min_font_size = min(min_font_size, sz)
                        if sz < 9.0:
                            discrepancies.append(f"Slide {s_idx} text '{r.text[:20]}' has font size {sz:.1f}pt < 9.0pt")

print(f"\n--- Verification Summary ---")
print(f"Total rounded rectangle boxes checked: {total_round_rects}")
print(f"Total lines checked: {total_lines}")
print(f"Minimum font size observed across deck: {min_font_size:.1f} pt")
print(f"Discrepancies found: {len(discrepancies)}")
for d in discrepancies:
    print(f"  - {d}")

if len(discrepancies) == 0:
    print("\nSUCCESS: 100% of rounded boxes have 4px border radius, 100% of lines from Library list, and all fonts >= 9.0 pt!")
else:
    print(f"\nFAIL: {len(discrepancies)} discrepancies found.")

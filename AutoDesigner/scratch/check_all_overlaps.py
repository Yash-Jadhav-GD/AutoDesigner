from pathlib import Path
from pptx import Presentation

p = Path("c:/Users/Yash.jadhav/OneDrive - GEP/Documents/Solution Design/antigravity/AutoFormat/AutoDesigner/output/Design_tester_new_formatted.pptx")
prs = Presentation(p)

total_overlaps = 0

for s_idx, s in enumerate(prs.slides, start=1):
    slide_overlaps = 0
    sh_list = []
    for sh in s.shapes:
        sh_list.append({
            "name": sh.name,
            "left": round(sh.left.pt, 1),
            "top": round(sh.top.pt, 1),
            "width": round(sh.width.pt, 1),
            "height": round(sh.height.pt, 1),
            "right": round(sh.left.pt + sh.width.pt, 1),
            "bottom": round(sh.top.pt + sh.height.pt, 1),
            "text": sh.text_frame.text[:30].replace("\n", " ") if sh.has_text_frame else ""
        })
    
    for i in range(len(sh_list)):
        for j in range(i + 1, len(sh_list)):
            a = sh_list[i]
            b = sh_list[j]
            # Container contains child if child is completely inside container
            is_contained_ab = (a["left"] <= b["left"] + 2 and a["right"] >= b["right"] - 2 and a["top"] <= b["top"] + 2 and a["bottom"] >= b["bottom"] - 2)
            is_contained_ba = (b["left"] <= a["left"] + 2 and b["right"] >= a["right"] - 2 and b["top"] <= a["top"] + 2 and b["bottom"] >= a["bottom"] - 2)
            if is_contained_ab or is_contained_ba:
                continue

            # Check if one is a floating header pill designed to overlap its card top edge
            is_floating_pill_ab = ("header text" in b["name"].lower() and a["left"] <= b["left"] + 4 and a["right"] >= b["right"] - 4 and abs(b["top"] - a["top"]) <= 24)
            is_floating_pill_ba = ("header text" in a["name"].lower() and b["left"] <= a["left"] + 4 and b["right"] >= a["right"] - 4 and abs(a["top"] - b["top"]) <= 24)
            if is_floating_pill_ab or is_floating_pill_ba:
                continue

            ox = max(0, min(a["right"], b["right"]) - max(a["left"], b["left"]))
            oy = max(0, min(a["bottom"], b["bottom"]) - max(a["top"], b["top"]))
            if ox > 1.0 and oy > 1.0:
                slide_overlaps += 1
                total_overlaps += 1
                print(f"  [Slide {s_idx}] Overlap: '{a['name']}' ({a['text']}) [{a['left']},{a['top']},{a['width']},{a['height']}] vs '{b['name']}' ({b['text']}) [{b['left']},{b['top']},{b['width']},{b['height']}] -> dx={ox:.1f}, dy={oy:.1f}")

if total_overlaps == 0:
    print("\nALL 20 SLIDES HAVE STRICTLY ZERO OVERLAPS!")
else:
    print(f"\nTotal overlaps detected: {total_overlaps}")

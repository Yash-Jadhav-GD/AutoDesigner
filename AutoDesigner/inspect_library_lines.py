import os
import pptx
from pptx.enum.shapes import MSO_SHAPE

lib_path = os.path.join(os.path.dirname(__file__), "Library_copy.pptx")
prs = pptx.Presentation(lib_path)
print(f"Total slides in Library_copy.pptx: {len(prs.slides)}")

target_names = [
    "Grey Centre Fade V",
    "Grey Solid V",
    "Grey Solid H",
    "Grey Centre Fade H",
    "Grey Side Fade R"
]

for s_idx, slide in enumerate(prs.slides, 1):
    found = []
    for sh in slide.shapes:
        if any(t.lower() in sh.name.lower() for t in target_names):
            found.append((sh.name, sh.shape_type, sh.width, sh.height, sh._element.tag))
    if found:
        print(f"\n--- Slide {s_idx} ---")
        for f in found:
            print(f"  Shape: {f[0]} | Type: {f[1]} | Size: {f[2]}x{f[3]} | Tag: {f[4]}")

# Also inspect all shape names on Slide 6 specifically
print("\n--- All shapes on Slide 6 ---")
s6 = prs.slides[5] # 0-indexed 5 is slide 6
for sh in s6.shapes:
    print(f"  '{sh.name}' | shape_type: {sh.shape_type} | pos: ({sh.left}, {sh.top}) | size: ({sh.width}, {sh.height})")

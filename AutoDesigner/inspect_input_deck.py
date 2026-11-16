import pptx
from pathlib import Path

p = Path("AutoDesigner/Design_tester_new_inspect.pptx")
prs = pptx.Presentation(p)
print(f"=== {p} (Total Slides: {len(prs.slides)}) ===")
for idx, slide in enumerate(prs.slides, 1):
    texts = [sh.text.strip().replace("\n", " ") for sh in slide.shapes if sh.has_text_frame and sh.text.strip()]
    first_few = " | ".join(texts[:3]) if texts else "Empty"
    print(f"\n--- Slide {idx} ({len(slide.shapes)} shapes) ---")
    print(f"  Sample text: {first_few[:120]}")

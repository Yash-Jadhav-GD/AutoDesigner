import pptx
from pathlib import Path

pptx_files = list(Path(".").rglob("*.pptx"))
print(f"Found {len(pptx_files)} pptx files")

for p in pptx_files:
    try:
        prs = pptx.Presentation(p)
        for s_idx, slide in enumerate(prs.slides, 1):
            for sh in slide.shapes:
                if "shadow" in sh.name.lower() or "decor" in sh.name.lower():
                    print(f"{p} Slide {s_idx}: '{sh.name}'")
    except Exception as e:
        # print(f"Could not read {p}: {e}")
        pass

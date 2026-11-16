import pptx
from pathlib import Path

prs = pptx.Presentation("AutoDesigner/Design_tester_new_inspect.pptx")
print(f"Total slides in Design tester new.pptx: {len(prs.slides)}")

for idx, slide in enumerate(prs.slides):
    shapes = slide.shapes
    title = "(No Title)"
    text_samples = []
    has_table = False
    
    for sh in shapes:
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                txt = p.text.strip()
                if txt:
                    text_samples.append(txt)
        elif sh.has_table:
            has_table = True
            for r in sh.table.rows:
                for c in r.cells:
                    for p in c.text_frame.paragraphs:
                        t = p.text.strip()
                        if t:
                            text_samples.append(t)

    # Guess title
    if text_samples:
        title = text_samples[0]
    
    print(f"\n--- Slide {idx+1} ({len(shapes)} shapes, {len(text_samples)} text lines, Table={has_table}) ---")
    print(f"Title: {title[:80]}")
    for t in text_samples[1:8]:
        print(f"  - {t[:80]}")
    if len(text_samples) > 8:
        print(f"  ... (+{len(text_samples)-8} more lines)")

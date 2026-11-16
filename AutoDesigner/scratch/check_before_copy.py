import pptx

prs = pptx.Presentation("AutoDesigner/Before_copy.pptx")
print(f"Before_copy has {len(prs.slides)} slides.")

for i, s in enumerate(prs.slides):
    titles = [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if p.text.strip()]
    t = titles[0] if titles else "(No text)"
    print(f"Slide {i+1}: {t[:70]}")

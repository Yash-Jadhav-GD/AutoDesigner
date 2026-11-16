import pptx, sys
sys.stdout.reconfigure(encoding='utf-8')

prs = pptx.Presentation(r'C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\Design_tester_new_inspect.pptx')
s7 = prs.slides[6]

texts = []
for shp in s7.shapes:
    if shp.has_text_frame:
        for p in shp.text_frame.paragraphs:
            t = p.text.strip()
            if t: texts.append(t)
    elif shp.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.GROUP:
        for gs in shp.shapes:
            if gs.has_text_frame:
                for gp in gs.text_frame.paragraphs:
                    t = gp.text.strip()
                    if t: texts.append(t)

print(f"Total texts extracted from Slide 7: {len(texts)}")
assert len(texts) == 39
print("All 39 text elements found successfully!")

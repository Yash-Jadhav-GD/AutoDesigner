import pptx, sys
sys.stdout.reconfigure(encoding='utf-8')

prs = pptx.Presentation(r'C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\Design_tester_new_inspect.pptx')

print(f"Total slides: {len(prs.slides)}")
for i, s in enumerate(prs.slides):
    t_list = []
    for shp in s.shapes:
        if shp.has_text_frame:
            for p in shp.text_frame.paragraphs:
                t = p.text.strip()
                if t: t_list.append(t)
        elif shp.has_table:
            for r in shp.table.rows:
                for c in r.cells:
                    for p in c.text_frame.paragraphs:
                        t = p.text.strip()
                        if t: t_list.append(t)
        elif shp.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.GROUP:
            for gs in shp.shapes:
                if gs.has_text_frame:
                    for gp in gs.text_frame.paragraphs:
                        t = gp.text.strip()
                        if t: t_list.append(t)
    print(f"Slide {i+1}: {len(t_list)} text elements")
    if i >= 6:
        print(f"  All texts for Slide {i+1}:")
        for idx, t in enumerate(t_list):
            print(f"    [{idx:02d}] {repr(t)}")

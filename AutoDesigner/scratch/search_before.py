import pptx, sys
sys.stdout.reconfigure(encoding='utf-8')

prs = pptx.Presentation(r'C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\Before_copy.pptx')
for idx, s in enumerate(prs.slides):
    t_list = []
    for shp in s.shapes:
        if shp.has_text_frame:
            for p in shp.text_frame.paragraphs:
                t = p.text.strip()
                if t: t_list.append(t)
        elif shp.has_table:
            t_list.append(f'<TABLE {len(shp.table.rows)}x{len(shp.table.columns)}>')
    full_str = ' '.join(t_list)
    for kw in ['paas 2.0', 'tier 2/3 value-creation', 'collaborative mfg', 'packaging portfolio', 'span of control driven']:
        if kw in full_str.lower():
            print(f'Before_copy Slide {idx+1} matched "{kw}": {t_list[0][:60]}')

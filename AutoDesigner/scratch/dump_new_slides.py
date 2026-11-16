import pptx, sys
sys.stdout.reconfigure(encoding='utf-8')

prs = pptx.Presentation(r'C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\Design_tester_new_inspect.pptx')

for s_idx in [6, 7, 8]:
    slide = prs.slides[s_idx]
    print(f'====================================')
    print(f'SLIDE {s_idx+1}')
    print(f'====================================')
    for i, s in enumerate(slide.shapes):
        t_list = []
        if s.has_text_frame:
            for p in s.text_frame.paragraphs:
                t = p.text.strip()
                if t: t_list.append(t)
        elif s.has_table:
            t_list.append(f'<TABLE {len(s.table.rows)}x{len(s.table.columns)}>')
            for r_idx, r in enumerate(s.table.rows):
                row_t = [c.text.strip().replace('\n', ' ') for c in r.cells]
                t_list.append(f'R{r_idx}: ' + ' | '.join(row_t))
        elif s.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.GROUP:
            t_list.append(f'<GROUP {len(s.shapes)} shapes>')
            for gs in s.shapes:
                if gs.has_text_frame:
                    for gp in gs.text_frame.paragraphs:
                        if gp.text.strip(): t_list.append(f'  {gs.name}: {gp.text.strip()}')
        if t_list:
            print(f'Shape {i:02d} ({s.name}): {t_list}')

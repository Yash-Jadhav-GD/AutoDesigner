import pptx
import sys
from pathlib import Path

BASE_DIR = Path("AutoDesigner")
prs = pptx.Presentation(BASE_DIR / "Design_tester_new_inspect.pptx")
print(f"Presentation has {len(prs.slides)} slides")

# --- Slide 1 ---
shapes_by_x = {}
for shape in prs.slides[0].shapes:
    if shape.has_text_frame and shape.text_frame.text.strip():
        x_col = round((shape.left / 12700) / 150)
        shapes_by_x.setdefault(x_col, []).append(shape)
s1_data = []
for k in sorted(shapes_by_x.keys()):
    shs = sorted(shapes_by_x[k], key=lambda s: s.top)
    s1_data.append({'header': shs[0].text_frame.text.strip(), 'items': [s.text_frame.text.strip() for s in shs[1:]]})

# --- Slide 2 ---
s2 = prs.slides[1]
l_shs = sorted([sh for sh in s2.shapes if sh.has_text_frame and sh.text_frame.text.strip() and (sh.left/12700) < 600.0], key=lambda s: s.top)
r_shs = sorted([sh for sh in s2.shapes if sh.has_text_frame and sh.text_frame.text.strip() and (sh.left/12700) >= 600.0], key=lambda s: s.top)
s2_data = {
    'left_pairs': [{'subheader': l_shs[i].text_frame.text.strip(), 'body': l_shs[i+1].text_frame.text.strip()} for i in range(0, len(l_shs), 2)],
    'right_header': r_shs[0].text_frame.text.strip(),
    'right_items': [sh.text_frame.text.strip() for sh in r_shs[1:]]
}

# --- Slide 3 ---
s3 = prs.slides[2]
shs3 = sorted([sh for sh in s3.shapes if sh.has_text_frame and sh.text_frame.text.strip()], key=lambda s: s.top)
s3_data = [{'subheader': shs3[i].text_frame.text.strip(), 'body': shs3[i+1].text_frame.text.strip()} for i in range(0, len(shs3), 2)]

# --- Slide 4 ---
s4 = prs.slides[3]
title_s4 = [p.text.strip() for sh in s4.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if 'Case Study 3' in p.text][0]
top_hdr_s4 = [p.text.strip() for sh in s4.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if 'Accelerated Sourcing Program' in p.text][0]
scope_s4 = [
    [p.text.strip() for sh in s4.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if 'Spend Assessed' in p.text][0],
    [sh.text_frame.text.strip().replace('\x0b', ' ') for sh in s4.shapes if sh.has_text_frame if 'Categories:' in sh.text_frame.text][0],
    [p.text.strip() for sh in s4.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if 'Regions-in-scope' in p.text][0],
    [p.text.strip() for sh in s4.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if 'Opportunities Identified' in p.text][0],
]
pain_hdr_s4 = [p.text.strip() for sh in s4.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if 'Pain Points' in p.text][0]
pain_shs = sorted([sh for sh in s4.shapes if sh.has_text_frame and (sh.left/914400) < 3.0 and (sh.top/914400) > 2.5 and 'Pain Points' not in sh.text_frame.text and 'Case Study' not in sh.text_frame.text and 'Opportunity' not in sh.text_frame.text and 'Spend' not in sh.text_frame.text and 'Key Outcomes' not in sh.text_frame.text], key=lambda s: s.top)
pain_bullets_s4 = [sh.text_frame.text.strip() for sh in pain_shs]

app_hdr_s4 = [p.text.strip() for sh in s4.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if p.text.strip() == 'Program Approach'][0]
tb75 = [sh for sh in s4.shapes if 'Assess and Prioritize' in (sh.text_frame.text if sh.has_text_frame else '')][0]
tb75_lines = [p.text.strip() for p in tb75.text_frame.paragraphs if p.text.strip()]
app_sec1_title = tb75_lines[0]
app_sec1_bullets = tb75_lines[1:]

tb71 = [sh for sh in s4.shapes if 'Develop and Execute' in (sh.text_frame.text if sh.has_text_frame else '')][0]
tb71_lines = [p.text.strip() for p in tb71.text_frame.paragraphs if p.text.strip()]
app_sec2_title = tb71_lines[0]
app_sec2_bullets = tb71_lines[1:]

opps_hdr_s4 = [p.text.strip() for sh in s4.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if 'Key Opportunities' in p.text][0]
tb16 = [sh for sh in s4.shapes if 'IT contract harmonization' in (sh.text_frame.text if sh.has_text_frame else '')][0]
opps_items_s4 = [p.text.strip() for p in tb16.text_frame.paragraphs if p.text.strip()]

outcomes_hdr_s4 = [p.text.strip() for sh in s4.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if 'Key Outcomes' in p.text][0]
outcome_tiles = sorted([sh for sh in s4.shapes if sh.has_text_frame and (sh.top/914400) > 6.1 and 'Key Outcomes' not in sh.text_frame.text], key=lambda s: s.left)
outcomes_metrics_s4 = [sh.text_frame.text.strip() for sh in outcome_tiles]

# --- Slide 5 ---
def extract_table_slide(s, expected_rows, is_s7=False):
    title = [p.text.strip() for sh in s.shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if sh.top/914400 < 1.0 and p.text.strip()][0]
    headers = [sh.text_frame.text.strip() for sh in sorted([sh for sh in s.shapes if sh.has_text_frame and 1.0 <= sh.top/914400 <= 1.35], key=lambda x: x.left)]
    
    row_shapes = [sh for sh in s.shapes if sh.has_text_frame and sh.top/914400 >= 1.4]
    rows = []
    threshold = 0.35 if expected_rows == 6 else 0.6
    for sh in sorted(row_shapes, key=lambda x: x.top):
        top_in = sh.top / 914400
        matched = False
        for r in rows:
            if abs(r['avg_top'] - top_in) < threshold:
                r['shapes'].append(sh)
                r['avg_top'] = sum(s.top/914400 for s in r['shapes']) / len(r['shapes'])
                matched = True
                break
        if not matched:
            rows.append({'avg_top': top_in, 'shapes': [sh]})
            
    assert len(rows) == expected_rows, f"Expected {expected_rows} rows, got {len(rows)}"
    
    parsed_rows = []
    for r in rows:
        sorted_shs = sorted(r['shapes'], key=lambda x: x.left)
        if is_s7:
            # S7 format: Col 0 = Master Title, Col 1 = Budget/Assets, Col 2 = Interventions, Col 3 = Impact
            # sorted_shs:
            # x ~ 0.42: title
            # x ~ 2.94: budget/assets
            # x ~ 6.05: interventions
            # x ~ 11.18: impact
            name_sh = [s for s in sorted_shs if s.left/914400 < 2.0][0]
            budget_sh = [s for s in sorted_shs if 2.0 <= s.left/914400 < 5.0][0]
            interv_sh = [s for s in sorted_shs if 5.0 <= s.left/914400 < 10.0][0]
            impact_sh = [s for s in sorted_shs if s.left/914400 >= 10.0][0]
            
            # budget_assets might have soft break or newline
            b_text = budget_sh.text_frame.text.strip().replace('\x0b', '\n')
            b_lines = [line.strip() for line in b_text.split('\n') if line.strip()]
            
            parsed_rows.append({
                'name': name_sh.text_frame.text.strip(),
                'budget_assets': b_lines,
                'interventions': [p.text.strip() for p in interv_sh.text_frame.paragraphs if p.text.strip()],
                'impact': [p.text.strip() for p in impact_sh.text_frame.paragraphs if p.text.strip()],
            })
        else:
            # S5 and S6 format: Col 0 = Badge & Name, Col 1 = Spend & Suppliers, Col 2 = Interventions, Col 3 = Impact
            # Badge vs Name: Badge is Graphic 858 or smaller box around x=0.67
            col0_shs = [s for s in sorted_shs if s.left/914400 < 2.5]
            badge_txt = ""
            name_txt = ""
            for s in col0_shs:
                t = s.text_frame.text.strip()
                if len(t) <= 2 or "Graphic" in s.name:
                    badge_txt = t
                else:
                    name_txt = t
            spend_sh = [s for s in sorted_shs if 2.5 <= s.left/914400 < 5.0][0]
            interv_sh = [s for s in sorted_shs if 5.0 <= s.left/914400 < 10.0][0]
            impact_sh = [s for s in sorted_shs if s.left/914400 >= 10.0][0]
            
            spend_lines = [p.text.strip() for p in spend_sh.text_frame.paragraphs if p.text.strip()]
            interv_lines = [p.text.strip() for p in interv_sh.text_frame.paragraphs if p.text.strip()]
            impact_lines = [p.text.strip() for p in impact_sh.text_frame.paragraphs if p.text.strip()]
            
            parsed_rows.append({
                'badge': badge_txt,
                'name': name_txt,
                'spend_sup': spend_lines,
                'interventions': interv_lines,
                'impact': impact_lines,
            })
            
    return {
        'title': title,
        'headers': headers,
        'rows': parsed_rows
    }

s5_data = extract_table_slide(prs.slides[4], expected_rows=4)
s6_data = extract_table_slide(prs.slides[5], expected_rows=6)
s7_data = extract_table_slide(prs.slides[6], expected_rows=3, is_s7=True)

print("Extraction summary:")
print("S1 cols:", len(s1_data))
print("S2 left:", len(s2_data['left_pairs']), "right:", len(s2_data['right_items']))
print("S3 pairs:", len(s3_data))
print("S4 pain:", len(pain_bullets_s4), "app1:", len(app_sec1_bullets), "app2:", len(app_sec2_bullets), "opps:", len(opps_items_s4), "out:", len(outcomes_metrics_s4))
print("S5 rows:", len(s5_data['rows']), "headers:", s5_data['headers'])
print("S6 rows:", len(s6_data['rows']), "headers:", s6_data['headers'])
print("S7 rows:", len(s7_data['rows']), "headers:", s7_data['headers'])

import pptx, sys
sys.stdout.reconfigure(encoding='utf-8')

prs = pptx.Presentation(r'C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\Design_tester_new_inspect.pptx')

# Slide 8 Table
s8 = prs.slides[7]
tbl8 = None
for s in s8.shapes:
    if s.has_table: tbl8 = s.table

print('=== SLIDE 8 TABLE CELLS ===')
for r_idx, r in enumerate(tbl8.rows):
    row_vals = [c.text.strip().replace('\n', ' ') for c in r.cells]
    print(f'Row {r_idx:02d}: {row_vals}')

# Slide 9 Table
s9 = prs.slides[8]
tbl9 = None
for s in s9.shapes:
    if s.has_table: tbl9 = s.table

print('\n=== SLIDE 9 TABLE CELLS ===')
for r_idx, r in enumerate(tbl9.rows):
    row_vals = [c.text.strip().replace('\n', ' ') for c in r.cells]
    print(f'Row {r_idx:02d}: {row_vals}')

import pptx, sys
sys.stdout.reconfigure(encoding='utf-8')

prs = pptx.Presentation(r'C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\After_copy.pptx')
s17 = prs.slides[16] # 0-indexed slide 17
tbl = None
for s in s17.shapes:
    if s.has_table: tbl = s.table

print('Table dimensions:', len(tbl.rows), 'x', len(tbl.columns))
print('Table shape XML sample:')
print(tbl._tbl.xml[:800])

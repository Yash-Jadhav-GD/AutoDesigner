import pptx, sys
sys.stdout.reconfigure(encoding='utf-8')

prs = pptx.Presentation(r'C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\After_copy.pptx')
s17 = prs.slides[16]
tbl = [s.table for s in s17.shapes if s.has_table][0]

for r_idx in range(len(tbl.rows)):
    row = tbl.rows[r_idx]
    cell0 = row.cells[0]
    tcPr = cell0._tc.tcPr
    fill_xml = tcPr.xml if tcPr is not None else 'None'
    txt = ' | '.join(c.text.strip().replace('\n', ' ') for c in row.cells)
    print(f'Row {r_idx}: {txt[:40]} | tcPr: {fill_xml[:150]}')

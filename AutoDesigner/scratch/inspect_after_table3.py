import pptx, sys
sys.stdout.reconfigure(encoding='utf-8')

prs = pptx.Presentation(r'C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\After_copy.pptx')
s17 = prs.slides[16]
tbl = [s.table for s in s17.shapes if s.has_table][0]

print('Row 0 tcPr full:')
print(tbl.rows[0].cells[0]._tc.tcPr.xml)
print('\nRow 1 tcPr full:')
print(tbl.rows[1].cells[0]._tc.tcPr.xml)

import pptx, sys
sys.stdout.reconfigure(encoding='utf-8')

prs = pptx.Presentation(r'C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\Design_tester_new_inspect.pptx')
s9 = prs.slides[8]

title = None
tbl_shape = None
for shp in s9.shapes:
    if shp.has_table:
        tbl_shape = shp
    elif shp.has_text_frame:
        txt = shp.text_frame.text.strip()
        if 'Span of Control Driven' in txt:
            title = txt

table_data = []
for r in tbl_shape.table.rows:
    row_cells = []
    for c in r.cells:
        p_list = [p.text.strip() for p in c.text_frame.paragraphs if p.text.strip()]
        row_cells.append(p_list)
    table_data.append(row_cells)

print('Title:', title)
print('Table rows:', len(table_data))
for idx, r in enumerate(table_data):
    print(f'R{idx}: {r}')

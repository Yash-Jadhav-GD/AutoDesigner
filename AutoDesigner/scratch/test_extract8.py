import pptx, sys
sys.stdout.reconfigure(encoding='utf-8')

prs = pptx.Presentation(r'C:\Users\Yash.jadhav\OneDrive - GEP\Documents\Solution Design\antigravity\AutoFormat\AutoDesigner\Design_tester_new_inspect.pptx')
s8 = prs.slides[7]

title = None
note = None
footer = None
tbl_shape = None
for shp in s8.shapes:
    if shp.has_table:
        tbl_shape = shp
    elif shp.has_text_frame:
        txt = shp.text_frame.text.strip()
        if 'Spend and Value Based' in txt:
            title = txt
        elif 'Packaging sub-categories' in txt:
            footer = txt
        elif 'Considered €200 MN' in txt:
            note = txt
    elif shp.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.GROUP:
        for gs in shp.shapes:
            if gs.has_text_frame:
                gtxt = gs.text_frame.text.strip()
                if 'Considered €200 MN' in gtxt:
                    note = gtxt

print('Title:', title[:50] if title else None)
print('Note:', note[:50] if note else None)
print('Footer:', footer[:50] if footer else None)
print('Table rows:', len(tbl_shape.table.rows))

from pptx import Presentation
from pptx.util import Pt
from pptx.oxml import parse_xml

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
tb = slide.shapes.add_textbox(Pt(10), Pt(10), Pt(100), Pt(100))
p = tb.text_frame.paragraphs[0]
p.text = 'Hello'
p.font.name = 'Aptos'
p.font.size = Pt(11)

pPr = p._p.get_or_add_pPr()
pPr.set('marL', '171450')
pPr.set('indent', '-171450')

bu_font = parse_xml('<a:buFont xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" typeface="Arial"/>')
bu_char = parse_xml('<a:buChar xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" char="•"/>')

insert_idx = len(pPr)
for i, child in enumerate(pPr):
    if child.tag.endswith(('tabLst', 'defRPr', 'extLst')):
        insert_idx = i
        break
pPr.insert(insert_idx, bu_font)
pPr.insert(insert_idx + 1, bu_char)

print(p._p.xml)
prs.save('AutoDesigner/scratch/test_schema_bullet.pptx')
print("Successfully saved test_schema_bullet.pptx!")

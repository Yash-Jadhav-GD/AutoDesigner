import pptx
from pptx import Presentation
from pptx.util import Pt
from pptx.oxml import parse_xml
from pptx.enum.text import MSO_ANCHOR

prs = Presentation()
prs.slide_width = Pt(960)
prs.slide_height = Pt(540)
slide = prs.slides.add_slide(prs.slide_layouts[6])

tx_box = slide.shapes.add_textbox(Pt(100), Pt(100), Pt(400), Pt(200))
tf = tx_box.text_frame
tf.word_wrap = True
tf.vertical_anchor = MSO_ANCHOR.MIDDLE

p = tf.paragraphs[0]
p.text = 'This is native bullet line one that wraps around to demonstrate hanging indent alignment in PowerPoint.'
p.font.name = 'Aptos'
p.font.size = Pt(11)

p2 = tf.add_paragraph()
p2.text = 'This is native bullet line two which is shorter.'
p2.font.name = 'Aptos'
p2.font.size = Pt(11)

def set_native_bullet(p, char='•', font_name='Arial', mar_l_pt=14.0, indent_pt=-14.0):
    pPr = p._p.get_or_add_pPr()
    pPr.set('marL', str(int(mar_l_pt * 12700)))
    pPr.set('indent', str(int(indent_pt * 12700)))
    for child in list(pPr):
        if 'bu' in child.tag:
            pPr.remove(child)
    bu_font = parse_xml(f'<a:buFont xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" typeface="{font_name}"/>')
    bu_char = parse_xml(f'<a:buChar xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" char="{char}"/>')
    pPr.append(bu_font)
    pPr.append(bu_char)

set_native_bullet(p)
set_native_bullet(p2)

prs.save('AutoDesigner/scratch_test_bullet.pptx')
print('Saved scratch_test_bullet.pptx!')

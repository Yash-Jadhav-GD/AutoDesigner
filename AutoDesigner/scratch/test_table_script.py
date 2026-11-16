import pptx
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.oxml import parse_xml
import io

prs = pptx.Presentation()
prs.slide_width = Pt(960)
prs.slide_height = Pt(540)
blank = prs.slide_layouts[6]
slide = prs.slides.add_slide(blank)

# Add a test table
rows = 4
cols = 3
shape = slide.shapes.add_table(rows, cols, Pt(30), Pt(60), Pt(900), Pt(200))
tbl = shape.table

def style_cell(cell, bg_hex, text, is_bold=False, font_size_pt=10, text_color_rgb=RGBColor(0x22, 0x1E, 0x1F), align=None):
    cell.fill.solid()
    r = int(bg_hex[0:2], 16)
    g = int(bg_hex[2:4], 16)
    b = int(bg_hex[4:6], 16)
    cell.fill.fore_color.rgb = RGBColor(r, g, b)
    
    # borders
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.set('anchor', 'ctr')
    
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = 'Aptos'
    p.font.size = Pt(font_size_pt)
    p.font.bold = is_bold
    p.font.color.rgb = text_color_rgb
    if align:
        p.alignment = align

for c_idx in range(cols):
    style_cell(tbl.cell(0, c_idx), "48397E", f"Header {c_idx+1}", is_bold=True, font_size_pt=11, text_color_rgb=RGBColor(0xFF, 0xFF, 0xFF))

buf = io.BytesIO()
prs.save(buf)
print("Table created and saved successfully! Bytes:", len(buf.getvalue()))

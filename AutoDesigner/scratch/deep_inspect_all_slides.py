import pptx
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Pt
from pathlib import Path

prs = pptx.Presentation("AutoDesigner/Design_tester_new_inspect.pptx")
print(f"Presentation has {len(prs.slides)} slides. Width={prs.slide_width.pt}pt, Height={prs.slide_height.pt}pt\n")

for s_idx, slide in enumerate(prs.slides):
    print("=" * 80)
    print(f"SLIDE {s_idx + 1} (Shapes: {len(slide.shapes)})")
    print("=" * 80)
    
    for sh_idx, shape in enumerate(slide.shapes):
        pos_str = f"left={shape.left.pt:.1f}pt, top={shape.top.pt:.1f}pt, w={shape.width.pt:.1f}pt, h={shape.height.pt:.1f}pt"
        sh_type = str(shape.shape_type)
        
        # Check text
        paras = []
        if shape.has_text_frame:
            for p in shape.text_frame.paragraphs:
                txt = p.text.strip()
                if txt:
                    font_sizes = [f"{r.font.size.pt:.1f}pt" for r in p.runs if r.font.size]
                    bolds = [str(r.font.bold) for r in p.runs if r.font.bold is not None]
                    sz_info = font_sizes[0] if font_sizes else (f"{p.font.size.pt:.1f}pt" if p.font.size else "def")
                    b_info = "B" if "True" in bolds or p.font.bold else ""
                    paras.append(f"['{txt[:60]}' sz={sz_info} {b_info}]")
        
        # Check table
        tbl_info = ""
        if shape.has_table:
            t = shape.table
            tbl_info = f" Table({len(t.rows)}x{len(t.columns)})"
            cell_texts = []
            for r_i, r in enumerate(t.rows):
                row_str = []
                for c_i, c in enumerate(r.cells):
                    ct = " ".join(p.text.strip() for p in c.text_frame.paragraphs if p.text.strip())
                    if ct:
                        row_str.append(f"[{r_i},{c_i}]: '{ct[:30]}'")
                if row_str:
                    cell_texts.append(" | ".join(row_str))
            tbl_info += "\n    " + "\n    ".join(cell_texts[:5])
        
        print(f"  Shape {sh_idx + 1} [{shape.name}] type={sh_type} {pos_str}{tbl_info}")
        if paras:
            for p_str in paras:
                print(f"    P: {p_str}")

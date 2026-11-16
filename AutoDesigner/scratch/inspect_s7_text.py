import pptx
from pptx.dml.color import RGBColor

prs = pptx.Presentation("AutoDesigner/Design_tester_new_inspect.pptx")
slide = prs.slides[6]

for i, sp in enumerate(slide.shapes):
    if not sp.has_text_frame:
        continue
    for p_idx, p in enumerate(sp.text_frame.paragraphs):
        txt = p.text.strip()
        if not txt:
            continue
        p_info = []
        for r_idx, r in enumerate(p.runs):
            col_info = "None"
            if r.font.color:
                try:
                    if r.font.color.type == pptx.enum.dml.MSO_COLOR_TYPE.RGB:
                        col_info = f"RGB({r.font.color.rgb})"
                    elif r.font.color.type == pptx.enum.dml.MSO_COLOR_TYPE.THEME:
                        col_info = f"THEME({r.font.color.theme_color})"
                except Exception as e:
                    col_info = f"ERR({e})"
            sz = r.font.size.pt if r.font.size else "None"
            p_info.append(f"r{r_idx}='{r.text}'(sz={sz},b={r.font.bold},clr={col_info})")
        print(f"Shape {i} [{sp.name}] p{p_idx}: {' '.join(p_info)}")

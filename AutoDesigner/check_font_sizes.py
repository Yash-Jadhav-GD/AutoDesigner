import pptx
from pathlib import Path

prs = pptx.Presentation("AutoDesigner/output/Design_tester_new_formatted.pptx")
print(f"Checking font sizes across {len(prs.slides)} slides in formatted presentation:\n")

below_9 = []

for s_idx, slide in enumerate(prs.slides, 1):
    for sh in slide.shapes:
        if sh.has_text_frame:
            for p_idx, p in enumerate(sh.text_frame.paragraphs):
                p_sz = p.font.size.pt if p.font and p.font.size else None
                for r in p.runs:
                    r_sz = r.font.size.pt if r.font and r.font.size else p_sz
                    if r_sz is not None and r_sz < 9.0:
                        below_9.append((s_idx, sh.name, r.text[:30], r_sz))
        if sh.has_table:
            for r_idx, row in enumerate(sh.table.rows):
                for c_idx, cell in enumerate(row.cells):
                    for p in cell.text_frame.paragraphs:
                        p_sz = p.font.size.pt if p.font and p.font.size else None
                        for r in p.runs:
                            r_sz = r.font.size.pt if r.font and r.font.size else p_sz
                            if r_sz is not None and r_sz < 9.0:
                                below_9.append((s_idx, f"Table[{r_idx},{c_idx}]", r.text[:30], r_sz))

print(f"Total occurrences of font size < 9.0: {len(below_9)}")
for b in below_9:
    print(f"  Slide {b[0]} | Shape: {b[1]} | Text: '{b[2]}' | Size: {b[3]}pt")

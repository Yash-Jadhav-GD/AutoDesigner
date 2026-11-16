from pptx import Presentation

prs = Presentation("AutoDesigner/Design tester new.pptx")
with open("AutoDesigner/scratch/slides_8_20_text_dump.txt", "w", encoding="utf-8") as out:
    for i in range(7, len(prs.slides)):
        s = prs.slides[i]
        out.write(f"\n========================================\n=== SLIDE {i+1} ===\n========================================\n")
        for sh in s.shapes:
            if sh.has_text_frame:
                paras = [p.text.strip() for p in sh.text_frame.paragraphs if p.text.strip()]
                if paras:
                    out.write(f"  [{sh.name}] (left={round(sh.left.pt,1)}, top={round(sh.top.pt,1)}, w={round(sh.width.pt,1)}, h={round(sh.height.pt,1)})\n")
                    for p_idx, p_text in enumerate(paras):
                        out.write(f"      p{p_idx}: {p_text}\n")
            elif sh.has_table:
                out.write(f"  [TABLE {sh.name}] ({len(sh.table.rows)} rows x {len(sh.table.columns)} cols) at ({round(sh.left.pt,1)}, {round(sh.top.pt,1)})\n")
                for r_idx, row in enumerate(sh.table.rows):
                    row_vals = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
                    out.write(f"    Row {r_idx}: {' | '.join(row_vals)}\n")
            elif sh.shape_type == 13: # picture
                out.write(f"  [PICTURE {sh.name}] at ({round(sh.left.pt,1)}, {round(sh.top.pt,1)}, w={round(sh.width.pt,1)}, h={round(sh.height.pt,1)})\n")
print("Dump completed.")

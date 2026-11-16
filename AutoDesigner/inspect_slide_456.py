import pptx

prs = pptx.Presentation(r'AutoDesigner/Design_tester_new_inspect.pptx')

for slide_idx in [3, 4, 5]:
    s = prs.slides[slide_idx]
    print(f"\n==================== SLIDE {slide_idx+1} (shapes: {len(s.shapes)}) ====================")
    for i, sh in enumerate(s.shapes):
        pos_str = f"({sh.left.pt:.1f}, {sh.top.pt:.1f}, {sh.width.pt:.1f}, {sh.height.pt:.1f})"
        print(f"Shape {i}: {sh.name} type={sh.shape_type} pos={pos_str}")
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                txt = p.text.strip()
                if txt:
                    font_size = p.font.size.pt if p.font.size else "None"
                    print(f"    P: '{txt}' font={p.font.name} size={font_size} bold={p.font.bold}")
        elif sh.has_table:
            t = sh.table
            print(f"    TABLE {len(t.rows)}x{len(t.columns)}")
            for r_idx, r in enumerate(t.rows):
                row_vals = []
                for c_idx, c in enumerate(r.cells):
                    cell_txt = " | ".join(p.text.strip() for p in c.text_frame.paragraphs if p.text.strip())
                    row_vals.append(cell_txt)
                print(f"      Row {r_idx}: {row_vals}")

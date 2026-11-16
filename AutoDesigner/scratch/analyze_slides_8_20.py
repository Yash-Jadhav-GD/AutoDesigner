from pptx import Presentation

prs = Presentation("AutoDesigner/Design tester new.pptx")
print(f"Total slides: {len(prs.slides)}")

for idx in range(7, len(prs.slides)):
    s = prs.slides[idx]
    print(f"\n==========================================")
    print(f"=== SLIDE {idx+1} ===")
    print(f"==========================================")
    text_shapes = []
    tables = []
    for sh in s.shapes:
        if sh.has_text_frame:
            txt = sh.text_frame.text.strip()
            if txt:
                text_shapes.append((sh.name, round(sh.left.pt, 1), round(sh.top.pt, 1), round(sh.width.pt, 1), round(sh.height.pt, 1), txt))
        elif sh.has_table:
            tables.append(f"Table ({len(sh.table.rows)} rows x {len(sh.table.columns)} cols) at ({round(sh.left.pt, 1)}, {round(sh.top.pt, 1)})")

    print(f"Text shapes count: {len(text_shapes)}, Tables: {tables}")
    for name, l, t, w, h, txt in text_shapes[:8]:
        summary_txt = txt.replace("\n", " -- ")[:90]
        print(f"  [{name}] ({l}, {t}, {w}, {h}): {summary_txt}")

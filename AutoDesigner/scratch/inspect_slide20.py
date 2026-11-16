from pptx import Presentation

prs = Presentation("AutoDesigner/Design tester new.pptx")
s20 = prs.slides[19]
print("Slide 20 total shapes:", len(s20.shapes))
for idx, sh in enumerate(s20.shapes):
    sh_type = type(sh).__name__
    txt = ""
    if sh.has_text_frame:
        txt = sh.text_frame.text.strip().replace("\n", " ")[:60]
    elif sh.has_table:
        txt = f"Table {len(sh.table.rows)}x{len(sh.table.columns)}"
    print(f"{idx}: [{sh_type}] {sh.name} (type={sh.shape_type}) -> {txt}")
    if hasattr(sh, "shapes"):
        print(f"   Group has {len(sh.shapes)} subshapes:")
        for sub_idx, sub in enumerate(sh.shapes):
            sub_txt = sub.text_frame.text.strip().replace("\n", " ")[:60] if sub.has_text_frame else ""
            print(f"     sub {sub_idx}: {sub.name} -> {sub_txt}")

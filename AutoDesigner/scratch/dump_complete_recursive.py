from pptx import Presentation

def extract_texts_from_shape(sh, indent=""):
    results = []
    if sh.has_text_frame:
        paras = [p.text.strip() for p in sh.text_frame.paragraphs if p.text.strip()]
        if paras:
            results.append((sh.name, paras))
    elif sh.has_table:
        table_rows = []
        for r in sh.table.rows:
            table_rows.append([c.text.strip().replace('\n', ' ') for c in r.cells])
        results.append((sh.name, table_rows))
    elif hasattr(sh, "shapes"): # GroupShape
        for sub in sh.shapes:
            results.extend(extract_texts_from_shape(sub, indent + "  "))
    return results

prs = Presentation("AutoDesigner/Design tester new.pptx")
with open("AutoDesigner/scratch/slides_8_20_complete_text.txt", "w", encoding="utf-8") as out:
    for i in range(7, len(prs.slides)):
        s = prs.slides[i]
        out.write(f"\n========================================\n=== SLIDE {i+1} ===\n========================================\n")
        for sh in s.shapes:
            extracted = extract_texts_from_shape(sh)
            for name, content in extracted:
                if isinstance(content, list) and content and isinstance(content[0], list):
                    out.write(f"  [TABLE {name}]:\n")
                    for row in content:
                        out.write(f"      {' | '.join(row)}\n")
                elif isinstance(content, list):
                    out.write(f"  [{name}]:\n")
                    for p in content:
                        out.write(f"      - {p}\n")
print("Complete text extraction finished.")

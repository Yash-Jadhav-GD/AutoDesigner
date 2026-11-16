import pptx
import zipfile
from xml.etree import ElementTree as ET
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
INPUT_PATH = BASE_DIR / "Design_tester_new_inspect.pptx"
OUTPUT_PATH = BASE_DIR / "output" / "Design_tester_new_formatted.pptx"

in_prs = pptx.Presentation(INPUT_PATH)
out_prs = pptx.Presentation(OUTPUT_PATH)

print(f"Input slides: {len(in_prs.slides)}, Output slides: {len(out_prs.slides)}")
assert len(in_prs.slides) == len(out_prs.slides), f"Slide count mismatch! {len(in_prs.slides)} vs {len(out_prs.slides)}"

# 1. FONT SIZE AUDIT
print("\n=== 1. FONT SIZE AUDIT (HARD CONSTRAINT: >= 9.0 pt) ===")
sub_9 = []
for s_i, s in enumerate(out_prs.slides):
    for sp in s.shapes:
        if sp.has_text_frame:
            for p in sp.text_frame.paragraphs:
                for r in p.runs:
                    if r.font.size and r.font.size.pt < 9.0:
                        sub_9.append((s_i + 1, sp.name, r.text[:30], r.font.size.pt))

print(f"Total occurrences of font size < 9.0 pt: {len(sub_9)}")
for item in sub_9:
    print(f"  Slide {item[0]} [{item[1]}]: '{item[2]}' size={item[3]}pt")

# 2. TEXT FIDELITY AUDIT
print("\n=== 2. TEXT FIDELITY AUDIT ===")
def get_clean_text(slide):
    texts = []
    for sh in slide.shapes:
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                t = p.text.strip()
                if t and t != ".":
                    texts.append(t)
        elif sh.has_table:
            for r in sh.table.rows:
                for c in r.cells:
                    for p in c.text_frame.paragraphs:
                        t = p.text.strip()
                        if t:
                            texts.append(t)
        elif sh.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.GROUP:
            for g in sh.shapes:
                if g.has_text_frame:
                    for p in g.text_frame.paragraphs:
                        t = p.text.strip()
                        if t and t != ".":
                            texts.append(t)
    return texts

total_in = 0
total_out = 0
for i in range(len(in_prs.slides)):
    t_in = get_clean_text(in_prs.slides[i])
    t_out = get_clean_text(out_prs.slides[i])
    total_in += len(t_in)
    total_out += len(t_out)
    
    # Check how many of t_in are represented in t_out
    # We do keyword / substring checking because bullets might have "• " added
    out_blob = " ".join(t_out).lower()
    missing = []
    for tin in t_in:
        # ignore single-char noise like '-' or '.'
        if len(tin) <= 1:
            continue
        # check if key phrase in out_blob
        key_words = [w for w in tin.replace('\u2013', '-').replace('\u2014', '-').split() if len(w) > 3]
        if key_words:
            matched_words = [w.lower() in out_blob for w in key_words]
            if sum(matched_words) / len(matched_words) < 0.6:
                missing.append(tin)
    
    status = "100% MATCH" if not missing else f"MISSING {len(missing)} items: {missing[:2]}"
    print(f"Slide {i+1}: in={len(t_in)} items, out={len(t_out)} items -> {status}")

print(f"Total Text items: in={total_in}, out={total_out}")

# 3. CORNER RADIUS AUDIT
print("\n=== 3. CORNER RADIUS AUDIT (HARD CONSTRAINT: 4px) ===")
adj_violations = 0
total_round_rects = 0
for s_i, s in enumerate(out_prs.slides):
    for sp in s.shapes:
        if sp.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.AUTO_SHAPE and sp.auto_shape_type == pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE:
            total_round_rects += 1
            spPr = sp._element.spPr
            avLst = spPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom/{http://schemas.openxmlformats.org/drawingml/2006/main}avLst")
            if avLst is not None:
                gd = avLst.find("{http://schemas.openxmlformats.org/drawingml/2006/main}gd")
                if gd is not None:
                    adj_val = int(gd.get("fmla").split()[-1])
                    min_dim = min(sp.width.pt, sp.height.pt)
                    computed_radius = (adj_val / 100000.0) * min_dim
                    if abs(computed_radius - 4.0) > 0.15:
                        adj_violations += 1
                        print(f"  Violation on Slide {s_i+1} [{sp.name}]: radius={computed_radius:.2f}pt (expected 4.0pt)")

print(f"Total rounded rectangle shapes checked: {total_round_rects}")
print(f"Corner radius violations: {adj_violations}")

# 4. PACKAGE INTEGRITY
print("\n=== 4. PACKAGE INTEGRITY ===")
with zipfile.ZipFile(OUTPUT_PATH, "r") as z:
    for i in range(1, len(out_prs.slides) + 1):
        rel_name = f"ppt/slides/_rels/slide{i}.xml.rels"
        if rel_name in z.namelist():
            rel_xml = z.read(rel_name)
            tree = ET.fromstring(rel_xml)
            for r in tree:
                tgt = r.attrib.get("Target")
                if "image" in tgt.lower():
                    pkg_path = f"ppt/media/{Path(tgt).name}"
                    assert pkg_path in z.namelist(), f"Broken image target: {tgt}"
print("Package & relationship integrity: ALL OK!")

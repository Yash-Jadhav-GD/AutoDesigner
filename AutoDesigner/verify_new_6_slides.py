import pptx
import zipfile
from xml.etree import ElementTree as ET
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
INPUT_PATH = BASE_DIR / "Design_tester_new_inspect.pptx"
OUTPUT_PATH = BASE_DIR / "output" / "Design_tester_new_formatted.pptx"

def get_slide_text_list(slide):
    texts = []
    for sh in slide.shapes:
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                t = p.text.strip()
                if t:
                    texts.append(t)
        elif sh.has_table:
            for r in sh.table.rows:
                for c in r.cells:
                    for p in c.text_frame.paragraphs:
                        t = p.text.strip()
                        if t:
                            texts.append(t)
    return texts

in_prs = pptx.Presentation(INPUT_PATH)
out_prs = pptx.Presentation(OUTPUT_PATH)

print(f"Input slides: {len(in_prs.slides)}, Output slides: {len(out_prs.slides)}")
num_slides = min(6, len(in_prs.slides), len(out_prs.slides))

total_in = 0
total_out = 0
all_perfect = True

for i in range(num_slides):
    t_in = get_slide_text_list(in_prs.slides[i])
    t_out = get_slide_text_list(out_prs.slides[i])
    total_in += len(t_in)
    total_out += len(t_out)
    
    # Normalize bullet characters and soft breaks
    norm_in = sorted([t.replace('\x0b', ' ').replace('\u2013', '-').replace('\u2014', '-').replace('•', '').strip() for t in t_in])
    norm_out = sorted([t.replace('\x0b', ' ').replace('\u2013', '-').replace('\u2014', '-').replace('•', '').strip() for t in t_out])
    
    diff = set(norm_in) ^ set(norm_out)
    if not diff:
        print(f"Slide {i+1}: in={len(t_in)}, out={len(t_out)} -> PERFECT 100% MATCH")
    else:
        print(f"Slide {i+1}: in={len(t_in)}, out={len(t_out)} -> DIFF DETECTED:")
        print(f"  Missing in out: {set(norm_in) - set(norm_out)}")
        print(f"  Extra in out: {set(norm_out) - set(norm_in)}")
        all_perfect = False

print(f"\nTOTAL TEXT: in={total_in}, out={total_out}, All Perfect={all_perfect}")

# Check relationship integrity
with zipfile.ZipFile(OUTPUT_PATH, "r") as z:
    for i in range(1, len(out_prs.slides) + 1):
        rels_path = f"ppt/slides/_rels/slide{i}.xml.rels"
        if rels_path in z.namelist():
            tree = ET.fromstring(z.read(rels_path))
            missing = []
            for rel in tree:
                tgt = rel.attrib.get("Target", "")
                if tgt.startswith("../media/"):
                    fn = tgt.replace("../", "ppt/")
                    if fn not in z.namelist():
                        missing.append(fn)
            if missing:
                print(f"Slide {i} missing media: {missing}")
            else:
                print(f"Slide {i} rel integrity: OK")

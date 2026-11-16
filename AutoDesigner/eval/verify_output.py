"""
Output fidelity check (generalises scratch/verify_all_9.py) for any source/output deck pair:
  - slide count is identical
  - every source word appears on the matching output slide (multiset, case-insensitive)
  - every r:id / r:embed in each output slide resolves to a relationship

    python eval/verify_output.py <source.pptx> <output.pptx>
"""

import re
import sys
import zipfile
from collections import Counter
from pathlib import Path
from typing import Dict, List

BASE_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE_DIR))

from engine.slide_model import read_presentation
from engine.understand import load

R_ATTR = re.compile(rb'r:(?:id|embed|link|pict)="([^"]+)"')
REL_ID = re.compile(rb'Id="([^"]+)"')


def dangling_relationships(path: str) -> Dict[int, List[str]]:
    out = {}
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        slides = sorted((n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)),
                        key=lambda n: int(re.search(r"(\d+)", n.rsplit("/", 1)[1]).group(1)))
        for n in slides:
            idx = int(re.search(r"slide(\d+)\.xml", n).group(1))
            rels = n.replace("slides/", "slides/_rels/") + ".rels"
            known = set(REL_ID.findall(z.read(rels))) if rels in names else set()
            missing = sorted({r.decode() for r in R_ATTR.findall(z.read(n))} - {k for k in known})
            missing = [m for m in missing if m.encode() not in known]
            if missing:
                out[idx] = missing
    return out


def verify(source: str, output: str) -> Dict:
    src = read_presentation(load(source))
    out = read_presentation(load(output))
    report = {"slides_in": len(src), "slides_out": len(out), "slides": [], "ok": True}
    if len(src) != len(out):
        report["ok"] = False
    dangling = dangling_relationships(output)
    for i, s in enumerate(src):
        o = out[i] if i < len(out) else None
        sw, ow = Counter(s.words()), Counter(o.words() if o else [])
        lost = sorted((sw - ow).elements())
        row = {"slide": i + 1, "words_in": sum(sw.values()), "words_out": sum(ow.values()),
               "lost": lost, "dangling_rels": dangling.get(i + 1, [])}
        if lost or row["dangling_rels"]:
            report["ok"] = False
        report["slides"].append(row)
    return report


def print_report(r: Dict):
    print(f"Slides: in={r['slides_in']} out={r['slides_out']}")
    for s in r["slides"]:
        flag = "OK" if not s["lost"] and not s["dangling_rels"] else "FAIL"
        extra = ""
        if s["lost"]:
            extra += f" lost {len(s['lost'])}: {' '.join(s['lost'][:10])}"
        if s["dangling_rels"]:
            extra += f" dangling rels: {s['dangling_rels']}"
        print(f"  slide {s['slide']:>2}: words in={s['words_in']} out={s['words_out']}  {flag}{extra}")
    print("ALL OK" if r["ok"] else "FIDELITY CHECK FAILED")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    rep = verify(sys.argv[1], sys.argv[2])
    print_report(rep)
    sys.exit(0 if rep["ok"] else 1)

"""
Accuracy eval: runs a mapper over the gold decks and scores the spec it produces.

    python eval/run_eval.py --mapper rules
    python eval/run_eval.py --mapper baseline --save

Metrics per slide:
  archetype  - chosen archetype equals gold
  fields     - share of gold leaf values reproduced at the same path (0 if archetype differs)
  coverage   - share of source words present in the spec (must be 1.0: never drop content)
  valid      - spec data passes the archetype schema
"""

import argparse
import json
import re
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE_DIR))

from engine.understand import analyze_deck, load
from engine.builders.schemas import validate_data

GOLD = [
    ("Design_tester_new_inspect.pptx", "specs/design_tester_new.json"),
    ("input/sample_draft.pptx", "specs/sample_draft.json"),
]


def _norm(v):
    return re.sub(r"\s+", " ", v).strip() if isinstance(v, str) else v


def leaves(data, path=()):
    if isinstance(data, dict):
        for k, v in data.items():
            yield from leaves(v, path + (k,))
    elif isinstance(data, list):
        for i, v in enumerate(data):
            yield from leaves(v, path + (i,))
    else:
        yield path, _norm(data)


def field_accuracy(gold, pred) -> float:
    g = dict(leaves(gold))
    p = dict(leaves(pred))
    g = {k: v for k, v in g.items() if v not in (None, "")}
    if not g:
        return 1.0
    return sum(1 for k, v in g.items() if p.get(k) == v) / len(g)


def score_deck(deck, gold_path, mapper, use_cache=True):
    gold = json.loads((BASE_DIR / gold_path).read_text(encoding="utf-8"))
    spec = analyze_deck(load(str(BASE_DIR / deck)), mapper=mapper, log=lambda *a: None, use_cache=use_cache)
    rows = []
    for i, (g, p) in enumerate(zip(gold["slides"], spec["slides"]), start=1):
        arch_ok = g["archetype"] == p["archetype"]
        try:
            validate_data(p["archetype"], p["data"])
            valid = True
        except Exception:
            valid = False
        rows.append({
            "slide": i,
            "gold": g["archetype"],
            "pred": p["archetype"],
            "archetype": arch_ok,
            "fields": round(field_accuracy(g["data"], p["data"]) if arch_ok else 0.0, 3),
            "coverage": p["coverage"],
            "valid": valid,
            "confidence": p.get("confidence"),
            "missing": p.get("missing_words", [])[:8],
        })
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--mapper", default="rules", choices=["baseline", "rules", "llm"])
    ap.add_argument("--save", action="store_true", help="Write eval/results/<mapper>.json")
    ap.add_argument("--no-cache", action="store_true")
    args = ap.parse_args()

    all_rows = []
    for deck, gold in GOLD:
        rows = score_deck(deck, gold, args.mapper, use_cache=not args.no_cache)
        print(f"\n{deck}")
        print(f"  {'#':>2}  {'gold':22} {'pred':22} arch  fields  cover  valid")
        for r in rows:
            print(f"  {r['slide']:>2}  {r['gold']:22} {r['pred']:22} {'ok' if r['archetype'] else '--':4}  "
                  f"{r['fields']:.2f}    {r['coverage']:.2f}   {'ok' if r['valid'] else '--'}"
                  + (f"   missing: {' '.join(r['missing'])}" if r["missing"] else ""))
        all_rows += [dict(r, deck=deck) for r in rows]

    n = len(all_rows)
    summary = {
        "mapper": args.mapper,
        "slides": n,
        "archetype_accuracy": round(sum(r["archetype"] for r in all_rows) / n, 3),
        "field_accuracy": round(sum(r["fields"] for r in all_rows) / n, 3),
        "full_coverage_slides": sum(r["coverage"] >= 1.0 for r in all_rows),
        "mean_coverage": round(sum(r["coverage"] for r in all_rows) / n, 3),
        "schema_valid": sum(r["valid"] for r in all_rows),
    }
    print("\nSUMMARY", json.dumps(summary))
    if args.save:
        out = BASE_DIR / "eval" / "results" / f"{args.mapper}.json"
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps({"summary": summary, "slides": all_rows}, indent=2, ensure_ascii=False), encoding="utf-8")
        print(f"Saved {out}")


if __name__ == "__main__":
    main()

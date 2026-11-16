"""
Build a deck from a JSON spec, or write a spec for any draft deck.

    python build_from_spec.py --spec specs/design_tester_new.json --output output/deck.pptx
    python build_from_spec.py --extract "Design_tester_new_inspect.pptx" --spec output/specs/my_deck.json

specs/ holds hand-checked gold specs used by eval/run_eval.py; write new specs elsewhere.
"""

import argparse
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))

from engine.builders.deck import build_deck, extract_spec, load_spec, save_spec
from engine.builders.registry import ARCHETYPES


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--spec", required=True, help="Path to the JSON spec (read, or written with --extract)")
    p.add_argument("--extract", help="Analyze this draft deck (rules mapper) into --spec instead of building")
    p.add_argument("--output", help="Output .pptx path (build mode)")
    p.add_argument("--icons", action="store_true", help="Add keyword-matched SVG icons to card_grid rows")
    p.add_argument("--list-archetypes", action="store_true")
    args = p.parse_args()

    if args.list_archetypes:
        print("\n".join(sorted(ARCHETYPES)))
        return

    if args.extract:
        save_spec(extract_spec(args.extract), args.spec)
        print(f"Wrote spec: {args.spec}")
        return

    if not args.output:
        p.error("--output is required when building")
    data = build_deck(load_spec(args.spec), icons=args.icons)
    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_bytes(data)
    print(f"Wrote deck: {out}")


if __name__ == "__main__":
    main()

"""
Spec-driven deck builder.

Spec format (JSON):
{
  "slides": [
    {"archetype": "three_column_pills", "data": {...}},
    ...
  ]
}
"""

import io
import json
from pathlib import Path
from typing import Any, Dict

import pptx
from pptx.util import Pt

from pydantic import ValidationError

from .common import load_library, inject_svg_assets
from ..svg_icons import embed_svg_icons
from .generic import load_palette
from .schemas import validate_data
from .registry import ARCHETYPES, SVG_ICON_ARCHETYPES, LibraryAssets


def validate_spec(spec: Dict[str, Any]) -> None:
    slides = spec.get("slides")
    if not isinstance(slides, list) or not slides:
        raise ValueError("Spec must contain a non-empty 'slides' list.")
    for i, s in enumerate(slides, start=1):
        name = s.get("archetype")
        if name not in ARCHETYPES:
            raise ValueError(f"Slide {i}: unknown archetype {name!r}. Known: {', '.join(sorted(ARCHETYPES))}")
        if "data" not in s:
            raise ValueError(f"Slide {i}: missing 'data'.")
        try:
            validate_data(name, s["data"])
        except ValidationError as e:
            raise ValueError(f"Slide {i} ({name}): data does not match schema:\n{e}") from None


def extract_spec(input_path: str, mapper: str = "rules") -> Dict[str, Any]:
    """Builds a spec from any draft deck via engine.understand."""
    from ..understand import analyze_deck, load
    return analyze_deck(load(input_path), mapper=mapper)


def build_deck(spec: Dict[str, Any], palette: str = "gep_purple", icons: bool = False) -> bytes:
    """Renders a spec to .pptx bytes. `palette` applies to the generic card archetypes;
    `icons` adds keyword-matched SVG icons to card_grid rows (engine/svg_icons.py)."""
    validate_spec(spec)
    lib_prs, lib_s1, lib_s3, lib_lines = load_library()
    lib = LibraryAssets(s1=lib_s1, s3=lib_s3, lines=lib_lines, palette=load_palette(palette),
                        prs=lib_prs, icons=icons)

    out = pptx.Presentation()
    out.slide_width = Pt(960)
    out.slide_height = Pt(540)

    for s in spec["slides"]:
        ARCHETYPES[s["archetype"]](out, s["data"], lib)

    buf = io.BytesIO()
    out.save(buf)
    raw = buf.getvalue()

    targets = {i: SVG_ICON_ARCHETYPES[s["archetype"]]
               for i, s in enumerate(spec["slides"], start=1) if s["archetype"] in SVG_ICON_ARCHETYPES}
    if targets:
        raw = inject_svg_assets(raw, targets=targets)
    if lib.svg_icons:
        raw = embed_svg_icons(raw, lib.svg_icons)
    return raw


def load_spec(path: str) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_spec(spec: Dict[str, Any], path: str) -> None:
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(spec, f, indent=2, ensure_ascii=False)

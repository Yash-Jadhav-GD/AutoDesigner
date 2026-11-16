"""
Slide understanding: turns a draft deck into a spec ({"slides": [{archetype, data, ...}]}).

Mappers:
  baseline - the original ContentAnalyzer, mapped to generic card archetypes (for comparison)
  rules    - local, feature-scored archetype choice + schema filling from the slide model
  llm      - Claude maps the slide model to an archetype + data; falls back to rules
  auto     - llm when Claude credentials resolve, otherwise rules
"""

import io
from typing import Any, Dict, List, Optional

import pptx

from ..slide_model import read_presentation, coverage, SlideModel


def _baseline_slide(slide, model: SlideModel) -> Dict[str, Any]:
    from ..analyzer import ContentAnalyzer
    d = ContentAnalyzer.analyze_slide(slide, slide_index=model.index)
    cards = [{
        "header": c.header,
        "items": [{"subheader": i.subheader, "body": i.body} for i in c.items],
        "kpi_metric": c.kpi_metric,
        "kpi_label": c.kpi_label,
    } for c in d.cards] or [{"header": d.title or "", "items": []}]
    return {"archetype": "card_grid", "data": {"title": d.title, "subtitle": d.subtitle, "cards": cards},
            "confidence": 0.0, "reasons": ["baseline analyzer"]}


def load(path: str):
    with open(path, "rb") as f:
        return pptx.Presentation(io.BytesIO(f.read()))


def analyze_deck(prs, mapper: str = "auto", log=print, use_cache: bool = True) -> Dict[str, Any]:
    models = read_presentation(prs)
    slides: List[Dict[str, Any]] = []

    if mapper == "auto":
        from .llm import credentials_available
        mapper = "llm" if credentials_available() else "rules"

    for slide, model in zip(prs.slides, models):
        if mapper == "baseline":
            out = _baseline_slide(slide, model)
        elif mapper == "rules":
            from .rules import map_slide
            out = map_slide(model)
        elif mapper == "llm":
            from .llm import map_slide_llm
            out = map_slide_llm(model, log=log, use_cache=use_cache)
        else:
            raise ValueError(f"Unknown mapper {mapper!r}")
        cov = coverage(model, out["data"])
        out["coverage"] = round(cov["coverage"], 4) if not cov["missing"] else min(round(cov["coverage"], 4), 0.9999)
        if cov["missing"]:
            out["missing_words"] = cov["missing"][:40]
        slides.append(out)

    return {"mapper": mapper, "slides": slides}

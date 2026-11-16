"""
Archetype registry: maps a spec archetype name to a builder.
Every entry is called as builder(prs, data, lib) where `lib` is a LibraryAssets.
"""

from dataclasses import dataclass, field
from typing import Any, Callable, Dict, List

from . import archetypes as a
from .generic import build_card_slide, load_palette


@dataclass
class LibraryAssets:
    s1: Any
    s3: Any
    lines: Dict[str, Any]
    palette: Dict[str, str] = field(default_factory=load_palette)
    prs: Any = None                       # the Library presentation (generic archetypes read prototypes from it)
    icons: bool = False                   # card_grid rows get keyword-matched SVG icons
    svg_icons: Dict[int, List] = field(default_factory=dict)  # slide number -> icons to embed


ARCHETYPES: Dict[str, Callable] = {
    "three_column_pills":      lambda prs, d, lib: a.build_slide1(prs, d, lib.s1, lib.lines),
    "split_process_insights":  lambda prs, d, lib: a.build_slide2(prs, d, lib.s1, lib.s3, lib.lines),
    "process_rows":            lambda prs, d, lib: a.build_slide3(prs, d, lib.s3, lib.lines),
    "assessment_kpis":         lambda prs, d, lib: a.build_slide4(prs, d, lib.s1, lib.s3, lib.lines),
    "workflow_cards":          lambda prs, d, lib: a.build_slide5(prs, d, lib.s3, lib.lines),
    "persona_grid":            lambda prs, d, lib: a.build_slide6(prs, d, lib.s3, lib.lines),
    "org_structure":           lambda prs, d, lib: a.build_slide7(prs, d, lib.lines),
    "allocation_table":        lambda prs, d, lib: a.build_slide8(prs, d),
    "sizing_table":            lambda prs, d, lib: a.build_slide9(prs, d),
    # Generic archetypes rendered by LayoutEngine + SlideStyler from Library prototypes
    "card_grid":               lambda prs, d, lib: build_card_slide(prs, d, "card_grid", lib),
    "kpi_cards":               lambda prs, d, lib: build_card_slide(prs, d, "kpi_cards", lib),
    "subheader_list":          lambda prs, d, lib: build_card_slide(prs, d, "subheader_list", lib),
    "data_box":                lambda prs, d, lib: build_card_slide(prs, d, "data_box", lib),
    "crd_icn_v1":              lambda prs, d, lib: build_card_slide(prs, d, "crd_icn_v1", lib),
    "important_box":           lambda prs, d, lib: build_card_slide(prs, d, "important_box", lib),
    "imp_box_v1":              lambda prs, d, lib: build_card_slide(prs, d, "imp_box_v1", lib),
}

# One-line description per archetype, shown to Claude and in --list-archetypes.
DESCRIPTIONS: Dict[str, str] = {
    "three_column_pills": "2-4 parallel columns, each a short header over a list of statements (e.g. Problem / Bridge / Future).",
    "split_process_insights": "Left: list of subheader+body process steps. Right: a headed side panel of short insight statements.",
    "process_rows": "Full-width list of subheader+body rows, no side panel.",
    "assessment_kpis": "Title+subtitle; left subheader+body list; right top a headed checklist of deliverables; right bottom a headed list of KPI rows (big value, label, sub-label).",
    "workflow_cards": "Title+subtitle; rows each with a letter badge, agent title, description, In:/Out: lines and an autonomy %.",
    "persona_grid": "Title+subtitle; a lead badge; a grid of team cards (count, FTE label, team name, regions); a total badge.",
    "org_structure": "Org chart: top banner, row of leadership boxes (title, optional sub, FTE count, colour-coded type), three WL1 pools (title, role badge, description, cards or regions), two-entry colour legend.",
    "allocation_table": "Title, optional note above and footer below, one table (each cell a list of lines).",
    "sizing_table": "Title and one table (each cell a list of lines).",
    "card_grid": "Generic: title/subtitle and 1-4 cards, each a header over items (optional subheader + body).",
    "kpi_cards": "Generic: title/subtitle and cards that each lead with a metric (kpi_metric + kpi_label), optional items.",
    "subheader_list": "Generic: title/subtitle and one card listing subheader + body items.",
    "data_box": "Generic: title/subtitle and 1-4 data cards with an icon circle in the card header, divider, and body text (CRD-ICN-V1).",
    "crd_icn_v1": "Generic: title/subtitle and data cards with an icon circle in the card header (CRD-ICN-V1).",
    "important_box": "Generic: title/subtitle and high-contrast callout card with top header tab and white fading dividers (IMP-BOX-V1).",
    "imp_box_v1": "Generic: title/subtitle and high-contrast callout card with top header tab and white fading dividers (IMP-BOX-V1).",
}

# Archetypes that use SVG icons injected from After_copy.pptx (see common.inject_svg_assets),
# mapped to the relationship-id prefix their builder writes.
SVG_ICON_ARCHETYPES = {"split_process_insights": "rIdSvgS2_", "assessment_kpis": "rIdSvgS4_"}
